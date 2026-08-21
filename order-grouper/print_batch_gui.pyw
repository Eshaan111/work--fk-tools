from __future__ import annotations

import ctypes
import json
import logging
import queue
import re
import socket
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any

from print_zpl_batch import print_raw_zpl
from raw_print_capture import (
    DEFAULT_LABEL_DIRECTORY,
    DEFAULT_PDF_INBOX,
    DEFAULT_RAW_DIRECTORY,
    receive_connection,
    save_job,
)


CAPTURE_HOST = "127.0.0.1"
CAPTURE_PORT = 9100
PHYSICAL_PRINTER = "TSC TTP-244 Pro"
PILE_DIRECTORY = DEFAULT_RAW_DIRECTORY.parent / "piles"
LOG_FILE = DEFAULT_RAW_DIRECTORY.parent / "batch-printer.log"
LOGGER = logging.getLogger("flipkart_batch_printer")
COLOR_ORDER = {"ICE": 0, "BEIGE": 1, "BLACK": 2, "WHITE": 3}
FIT_ORDER = {"BAGGY": 0, "PLAIN": 1}
SIZE_ORDER = {26: 0, 28: 1, 30: 2, 32: 3, 34: 4, 36: 5}


def configure_runtime_logging() -> None:
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    if LOGGER.handlers:
        return
    handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s %(threadName)s %(message)s")
    )
    LOGGER.addHandler(handler)
    LOGGER.setLevel(logging.INFO)


def enable_high_dpi_rendering() -> None:
    """Let Windows render Tk at the monitor's native DPI instead of scaling it."""
    try:
        user32 = ctypes.windll.user32
        user32.SetProcessDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        user32.SetProcessDpiAwarenessContext.restype = ctypes.c_bool
        if user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except (AttributeError, OSError):
        pass

    try:
        if ctypes.windll.shcore.SetProcessDpiAwareness(2) == 0:
            return
    except (AttributeError, OSError):
        pass

    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except (AttributeError, OSError):
        pass


def category_heading(record: dict[str, object]) -> str:
    if bool(record.get("is_mix")):
        return "MIX"
    color = str(record["color"])
    size = int(record["size"])
    if color in {"ICE", "BEIGE"}:
        return f"{color}-{size}"
    return f"{color}-{record['fit']}-{size}"


def seller_name(record: dict[str, object]) -> str:
    normalized = str(record.get("seller") or "").strip().upper()
    if "PRABHU ENTERPRISES" in normalized:
        return "PRABHU"
    if "SEEMA ENTERPRISES" in normalized:
        return "SEEMA"
    cleaned = re.sub(r"[^A-Z0-9 &./_-]", "", normalized).strip()
    return cleaned or "UNKNOWN SELLER"


def build_print_summary(
    metadata: dict[str, object],
) -> list[dict[str, object]]:
    """Count captured order labels by color/fit and size for the UI summary."""
    labels_summary = Path(str(metadata["labels_summary"]))
    records = json.loads(labels_summary.read_text(encoding="utf-8"))
    grouped: dict[tuple[str, str], dict[int, int]] = {}
    mix_count = 0

    for record in records:
        if bool(record.get("is_mix")):
            mix_count += 1
            continue
        try:
            color = str(record["color"])
            fit = str(record["fit"])
            size = int(record["size"])
            if color not in COLOR_ORDER or fit not in FIT_ORDER or size not in SIZE_ORDER:
                raise ValueError("unrecognized sorting value")
        except (KeyError, TypeError, ValueError):
            mix_count += 1
            continue
        sizes = grouped.setdefault((color, fit), {})
        sizes[size] = sizes.get(size, 0) + 1

    summary: list[dict[str, object]] = []
    for color, fit in sorted(
        grouped,
        key=lambda value: (COLOR_ORDER[value[0]], FIT_ORDER[value[1]]),
    ):
        sizes = grouped[(color, fit)]
        summary.append(
            {
                "heading": f"{color}-{fit}",
                "total": sum(sizes.values()),
                "sizes": [
                    {"size": size, "count": sizes[size]}
                    for size in sorted(sizes, key=SIZE_ORDER.__getitem__)
                ],
            }
        )
    if mix_count:
        summary.append({"heading": "MIX", "total": mix_count, "sizes": []})
    return summary


def make_category_label(heading: str, label_count: int, sellers: str) -> bytes:
    """Create a standalone divider label; captured labels are never edited."""
    font_width = max(54, min(96, 900 // len(heading)))
    seller_width = max(42, min(92, 860 // len(sellers)))
    zpl = (
        "\r\n^XA\r\n"
        "^CI28\r\n"
        "^PW599\r\n"
        "^LL991\r\n"
        "^LH0,0\r\n"
        "^FO20,245\r\n"
        f"^A0N,115,{font_width}\r\n"
        "^FB559,1,0,C,0\r\n"
        f"^FD{heading}^FS\r\n"
        "^FO20,445\r\n"
        "^A0N,82,70\r\n"
        "^FB559,1,0,C,0\r\n"
        f"^FDCOUNT: {label_count}^FS\r\n"
        "^FO20,585\r\n"
        f"^A0N,92,{seller_width}\r\n"
        "^FB559,1,0,C,0\r\n"
        f"^FD{sellers}^FS\r\n"
        "^PQ1\r\n"
        "^XZ\r\n"
    )
    return zpl.encode("ascii")


def apply_sorting(source: Path, metadata: dict[str, object]) -> Path:
    """Reorder complete ZPL print units while preserving their original bytes."""
    labels_summary = Path(str(metadata["labels_summary"]))
    records = json.loads(labels_summary.read_text(encoding="utf-8"))
    data = source.read_bytes()
    required_sort_fields = {"color", "fit", "size", "is_mix", "sort_tag"}
    for index, record in enumerate(records, start=1):
        missing = required_sort_fields.difference(record)
        if missing:
            raise RuntimeError(
                f"Label {index} is missing sorting data: "
                f"{', '.join(sorted(missing))}. Re-extract this legacy batch first."
            )

    all_blocks = list(re.finditer(rb"\^XA.*?\^XZ", data, re.I | re.S))
    printable_blocks = [
        match
        for match in all_blocks
        if re.search(rb"\^(?:FD|GFA|BC|B[CQXR]|BX)", match.group(), re.I)
    ]
    if len(printable_blocks) != len(records):
        raise RuntimeError(
            "Cannot safely sort this batch: extracted label count does not match "
            "the printable ZPL block count."
        )

    # Each unit includes any exact printer-setup bytes immediately before its
    # printable label. The units are only permuted; their contents are untouched.
    units: list[bytes] = []
    cursor = 0
    for match in printable_blocks:
        units.append(data[cursor : match.end()])
        cursor = match.end()
    trailing_bytes = data[cursor:]

    def sort_key(item: tuple[int, dict[str, object]]) -> tuple[int, int, int, int, int]:
        original_index, record = item
        if bool(record.get("is_mix")):
            return (1, 0, 0, 0, original_index)
        return (
            0,
            COLOR_ORDER[str(record["color"])],
            FIT_ORDER[str(record["fit"])],
            SIZE_ORDER[int(record["size"])],
            original_index,
        )

    ordered = sorted(enumerate(records), key=sort_key)
    category_stats: dict[str, dict[str, object]] = {}
    for _, record in ordered:
        heading = category_heading(record)
        stats = category_stats.setdefault(heading, {"count": 0, "sellers": []})
        stats["count"] = int(stats["count"]) + 1
        alias = seller_name(record)
        seller_list = stats["sellers"]
        if isinstance(seller_list, list) and alias not in seller_list:
            seller_list.append(alias)

    sorted_directory = source.parent.parent / "sorted"
    sorted_directory.mkdir(parents=True, exist_ok=True)
    sorted_source = sorted_directory / source.name

    output_chunks: list[bytes] = []
    manifest: list[dict[str, object]] = []
    previous_heading: str | None = None
    print_position = 0
    separator_count = 0
    for original_index, record in ordered:
        heading = category_heading(record)
        if heading != previous_heading:
            stats = category_stats[heading]
            sellers = " / ".join(str(value) for value in stats["sellers"])
            category_count = int(stats["count"])
            output_chunks.append(
                make_category_label(heading, category_count, sellers)
            )
            print_position += 1
            separator_count += 1
            manifest.append(
                {
                    "print_position": print_position,
                    "kind": "category_separator",
                    "separator_text": heading,
                    "category_label_count": category_count,
                    "sellers": sellers,
                    "sort_tag": record.get("sort_tag"),
                }
            )
            previous_heading = heading

        output_chunks.append(units[original_index])
        print_position += 1
        manifest.append(
            {
                "print_position": print_position,
                "kind": "flipkart_label",
                "original_label_number": original_index + 1,
                "order_id": record.get("order_id"),
                "sku": record.get("sku"),
                "sort_tag": record.get("sort_tag"),
                "is_mix": record.get("is_mix"),
            }
        )
    sorted_source.write_bytes(b"".join(output_chunks) + trailing_bytes)

    manifest_path = labels_summary.parent / "sort-order.json"
    manifest_path.write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    metadata["sorting_applied"] = True
    metadata["sorted_file"] = str(sorted_source)
    metadata["sort_manifest"] = str(manifest_path)
    metadata["separator_count"] = separator_count
    metadata["printable_count"] = len(records) + separator_count
    source.with_suffix(source.suffix + ".json").write_text(
        json.dumps(metadata, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return sorted_source


class BatchCaptureWorker(threading.Thread):
    def __init__(
        self,
        events: queue.Queue[dict[str, Any]],
        pile: list[tuple[Path, dict[str, object]]] | None = None,
    ) -> None:
        super().__init__(daemon=True)
        self.events = events
        self.decisions: queue.Queue[str] = queue.Queue()
        self.commands: queue.Queue[tuple[str, str | None]] = queue.Queue()
        self.pile = list(pile or [])
        self.stop_event = threading.Event()
        self.server: socket.socket | None = None

    def emit(self, event_type: str, **details: object) -> None:
        self.events.put({"type": event_type, **details})

    def stop(self) -> None:
        self.stop_event.set()
        if self.server is not None:
            try:
                self.server.close()
            except OSError:
                pass

    def submit_decision(self, decision: str) -> None:
        self.decisions.put(decision)

    def request_print_pile(self, mode: str) -> None:
        self.commands.put(("print_pile", mode))

    def pile_totals(self) -> tuple[int, int]:
        return (
            len(self.pile),
            sum(int(metadata.get("label_count", 0)) for _, metadata in self.pile),
        )

    def emit_pile_updated(self) -> None:
        job_count, label_count = self.pile_totals()
        self.emit("pile_updated", job_count=job_count, label_count=label_count)

    def wait_for_decision(self) -> str:
        while not self.stop_event.is_set():
            try:
                return self.decisions.get(timeout=0.5)
            except queue.Empty:
                continue
        return "cancel"

    def prepare_pile_job(self) -> tuple[Path, dict[str, object]]:
        pile_id = f"pile-{datetime.now():%Y%m%d-%H%M%S-%f}"
        pile_directory = PILE_DIRECTORY / pile_id
        pile_directory.mkdir(parents=True, exist_ok=False)
        pile_source = pile_directory / f"{pile_id}.zpl"
        labels_summary = pile_directory / "labels.json"

        combined_data = bytearray()
        combined_records: list[dict[str, object]] = []
        for source, metadata in self.pile:
            combined_data.extend(source.read_bytes())
            summary_path = Path(str(metadata["labels_summary"]))
            records = json.loads(summary_path.read_text(encoding="utf-8"))
            combined_records.extend(records)

        pile_source.write_bytes(bytes(combined_data))
        labels_summary.write_text(
            json.dumps(combined_records, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        metadata: dict[str, object] = {
            "job_id": pile_id,
            "format": "zpl",
            "label_count": len(combined_records),
            "labels_summary": str(labels_summary),
            "pile_job_count": len(self.pile),
            "raw_file": str(pile_source),
        }
        pile_source.with_suffix(".zpl.json").write_text(
            json.dumps(metadata, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        return pile_source, metadata

    def print_saved_job(
        self,
        source: Path,
        metadata: dict[str, object],
        mode: str,
        context: str,
    ) -> None:
        order_label_count = int(metadata.get("label_count", 0))
        printable_source = source
        separator_count = 0
        printable_count = order_label_count
        if mode == "sorted":
            self.emit("sorting", label_count=order_label_count, context=context)
            printable_source = apply_sorting(source, metadata)
            separator_count = int(metadata.get("separator_count", 0))
            printable_count = int(
                metadata.get("printable_count", order_label_count)
            )

        self.emit(
            "printing",
            label_count=printable_count,
            order_label_count=order_label_count,
            separator_count=separator_count,
            context=context,
            mode=mode,
        )
        windows_job_id, verified_count = print_raw_zpl(
            printable_source,
            PHYSICAL_PRINTER,
        )
        summary: list[dict[str, object]] = []
        if mode == "sorted" or context == "pile":
            try:
                summary = build_print_summary(metadata)
            except Exception:
                LOGGER.exception("Could not build the printed-label summary")
        self.emit(
            "printed",
            source=str(printable_source),
            label_count=verified_count,
            order_label_count=order_label_count,
            windows_job_id=windows_job_id,
            context=context,
            mode=mode,
            summary=summary,
        )

    def process_commands(self) -> None:
        while True:
            try:
                command, argument = self.commands.get_nowait()
            except queue.Empty:
                return

            if command != "print_pile" or argument not in {"sorted", "original"}:
                continue
            if not self.pile:
                self.emit("pile_empty")
                continue
            try:
                source, metadata = self.prepare_pile_job()
                self.print_saved_job(source, metadata, argument, "pile")
                self.pile.clear()
                self.emit_pile_updated()
            except Exception as error:
                LOGGER.exception("Pile printing failed")
                self.emit(
                    "batch_error",
                    message=f"{type(error).__name__}: {error}",
                    context="pile",
                )

    def run(self) -> None:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as server:
                self.server = server
                server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                server.bind((CAPTURE_HOST, CAPTURE_PORT))
                server.listen(8)
                server.settimeout(0.5)
                self.emit("listening")

                while not self.stop_event.is_set():
                    self.process_commands()
                    try:
                        connection, peer = server.accept()
                    except socket.timeout:
                        continue
                    except OSError:
                        if self.stop_event.is_set():
                            break
                        raise

                    with connection:
                        self.process_connection(connection, peer)
        except Exception as error:
            LOGGER.exception("Capture worker stopped unexpectedly")
            self.emit("fatal", message=f"{type(error).__name__}: {error}")
        finally:
            self.server = None
            self.emit("stopped")

    def process_connection(
        self,
        connection: socket.socket,
        peer: tuple[str, int],
    ) -> None:
        try:
            data = receive_connection(
                connection,
                idle_timeout=2.0,
                maximum_bytes=100 * 1024 * 1024,
            )
            if not data:
                self.emit("empty_connection")
                return

            self.emit("capturing", byte_count=len(data))
            source, metadata = save_job(
                data,
                raw_directory=DEFAULT_RAW_DIRECTORY,
                pdf_inbox=DEFAULT_PDF_INBOX,
                labels_directory=DEFAULT_LABEL_DIRECTORY,
                peer=peer,
            )
            job_format = str(metadata.get("format", "unknown"))
            label_count = int(metadata.get("label_count", 0))
            self.emit(
                "captured",
                source=str(source),
                job_format=job_format,
                label_count=label_count,
            )

            if job_format != "zpl":
                raise RuntimeError(
                    f"Captured {job_format} data; automatic printing only accepts ZPL."
                )
            if label_count < 1:
                processing_error = metadata.get("label_processing_error")
                raise RuntimeError(
                    str(processing_error or "No printable labels were detected.")
                )

            self.emit(
                "decision_required",
                label_count=label_count,
                source=str(source),
            )
            decision = self.wait_for_decision()
            if decision == "sorted":
                self.print_saved_job(source, metadata, "sorted", "batch")
            elif decision == "original":
                self.print_saved_job(source, metadata, "original", "batch")
            elif decision == "pile":
                self.pile.append((source, metadata))
                self.emit_pile_updated()
                self.emit("added_to_pile", label_count=label_count)
            else:
                self.emit("batch_cancelled", label_count=label_count)
        except Exception as error:
            LOGGER.exception("Incoming batch processing failed")
            self.emit(
                "batch_error",
                message=f"{type(error).__name__}: {error}",
                context="batch",
            )


class BatchPrinterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.events: queue.Queue[dict[str, Any]] = queue.Queue()
        self.worker: BatchCaptureWorker | None = None
        root.report_callback_exception = self.report_callback_exception

        root.title("Flipkart Batch Printer")
        display_scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
        root.geometry(
            f"{round(800 * display_scale)}x{round(560 * display_scale)}"
        )
        root.minsize(
            round(720 * display_scale),
            round(500 * display_scale),
        )
        root.configure(background="#F2F7F3")
        root.protocol("WM_DELETE_WINDOW", self.close)

        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("App.TFrame", background="#F2F7F3")
        style.configure("Header.TFrame", background="#166534")
        style.configure("HeaderTitle.TLabel", background="#166534", foreground="#FFFFFF", font=("Segoe UI", 22, "bold"))
        style.configure("HeaderSub.TLabel", background="#166534", foreground="#D1FAE5", font=("Segoe UI", 10))
        style.configure("Card.TFrame", background="#FFFFFF", bordercolor="#DDE8E0", borderwidth=1, relief="solid")
        style.configure("Eyebrow.TLabel", background="#FFFFFF", foreground="#4B6353", font=("Segoe UI", 9, "bold"))
        style.configure("Status.TLabel", background="#FFFFFF", foreground="#153D25", font=("Segoe UI", 16, "bold"))
        style.configure("Detail.TLabel", background="#FFFFFF", foreground="#587061", font=("Segoe UI", 10))
        style.configure("Footer.TLabel", background="#F2F7F3", foreground="#587061", font=("Segoe UI", 9))
        style.configure("Primary.TButton", background="#15803D", foreground="#FFFFFF", borderwidth=0, padding=(18, 10), font=("Segoe UI", 10, "bold"))
        style.map("Primary.TButton", background=[("active", "#166534"), ("disabled", "#A7C7B1")], foreground=[("disabled", "#F3F7F4")])
        style.configure("Secondary.TButton", background="#E8F5EC", foreground="#166534", bordercolor="#B9D9C3", padding=(18, 10), font=("Segoe UI", 10, "bold"))
        style.map("Secondary.TButton", background=[("active", "#D8EDDF")])
        style.configure("Listening.TLabel", background="#DCFCE7", foreground="#166534", padding=(10, 5), font=("Segoe UI", 9, "bold"))
        style.configure("Stopped.TLabel", background="#EEF2EF", foreground="#66756B", padding=(10, 5), font=("Segoe UI", 9, "bold"))
        style.configure("Error.TLabel", background="#FEE2E2", foreground="#991B1B", padding=(10, 5), font=("Segoe UI", 9, "bold"))
        style.configure("Treeview", background="#FFFFFF", fieldbackground="#FFFFFF", foreground="#294634", borderwidth=0, rowheight=30, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", background="#EAF4ED", foreground="#28553A", borderwidth=0, padding=(8, 8), font=("Segoe UI", 9, "bold"))
        style.map("Treeview.Heading", background=[("active", "#DCEDE1")])

        self.status = tk.StringVar(value="Starting capture service...")
        self.detail = tk.StringVar(value="")
        self.listener_state = tk.StringVar(value="STOPPED")
        self.pile_state = tk.StringVar(value="Pile: 0 jobs / 0 labels")
        self.pile_job_count = 0
        self.pile_label_count = 0

        container = ttk.Frame(root, style="App.TFrame")
        container.pack(fill="both", expand=True)

        header = ttk.Frame(container, style="Header.TFrame", padding=(24, 18))
        header.pack(fill="x")
        ttk.Label(
            header,
            text="Flipkart Batch Printer",
            style="HeaderTitle.TLabel",
        ).pack(anchor="w")
        ttk.Label(
            header,
            text=f"Capture {CAPTURE_HOST}:{CAPTURE_PORT}  |  Printer {PHYSICAL_PRINTER}",
            style="HeaderSub.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        body = ttk.Frame(container, style="App.TFrame", padding=(22, 18, 22, 20))
        body.pack(fill="both", expand=True)

        status_frame = ttk.Frame(body, style="Card.TFrame", padding=18)
        status_frame.pack(fill="x")
        status_top = ttk.Frame(status_frame, style="Card.TFrame")
        status_top.pack(fill="x")
        ttk.Label(status_top, text="CURRENT STATUS", style="Eyebrow.TLabel").pack(side="left")
        self.listener_badge = ttk.Label(status_top, textvariable=self.listener_state, style="Stopped.TLabel")
        self.listener_badge.pack(side="right")
        ttk.Label(
            status_frame,
            textvariable=self.status,
            style="Status.TLabel",
        ).pack(anchor="w", pady=(12, 0))
        ttk.Label(
            status_frame,
            textvariable=self.detail,
            style="Detail.TLabel",
            wraplength=700,
        ).pack(anchor="w", pady=(5, 0))

        controls = ttk.Frame(status_frame, style="Card.TFrame")
        controls.pack(fill="x", pady=(16, 0))
        self.start_button = ttk.Button(
            controls,
            text="Start listening",
            command=self.start,
            style="Primary.TButton",
        )
        self.start_button.pack(side="left")
        self.stop_button = ttk.Button(
            controls,
            text="Stop listening",
            command=self.stop,
            state="disabled",
            style="Secondary.TButton",
        )
        self.stop_button.pack(side="left", padx=(8, 0))
        self.print_pile_button = ttk.Button(
            controls,
            text="Print pile",
            command=self.print_pile,
            state="disabled",
            style="Primary.TButton",
        )
        self.print_pile_button.pack(side="right")
        ttk.Label(
            controls,
            textvariable=self.pile_state,
            style="Detail.TLabel",
        ).pack(side="right", padx=(0, 12))

        history_frame = ttk.Frame(body, style="Card.TFrame", padding=14)
        history_frame.pack(fill="both", expand=True, pady=(14, 0))
        ttk.Label(history_frame, text="RECENT BATCHES", style="Eyebrow.TLabel").pack(anchor="w", pady=(0, 10))
        self.history = ttk.Treeview(
            history_frame,
            columns=("time", "labels", "job", "result"),
            show="headings",
            height=6,
        )
        self.history.heading("time", text="Time")
        self.history.heading("labels", text="Labels")
        self.history.heading("job", text="Windows job")
        self.history.heading("result", text="Result")
        self.history.column("time", width=90, anchor="center")
        self.history.column("labels", width=75, anchor="center")
        self.history.column("job", width=100, anchor="center")
        self.history.column("result", width=380)
        self.history.pack(fill="both", expand=True)
        self.history.tag_configure("success", foreground="#166534")
        self.history.tag_configure("error", foreground="#991B1B")

        ttk.Label(
            body,
            text=(
                "SORT ORDER   Color  >  BAGGY / PLAIN  >  Size   |   "
                "Divider before each group   |   MIX last"
            ),
            style="Footer.TLabel",
        ).pack(anchor="w", pady=(12, 0))

        root.after(100, self.start)
        root.after(100, self.poll_events)

    def start(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            return
        retained_pile = list(self.worker.pile) if self.worker is not None else []
        self.worker = BatchCaptureWorker(self.events, pile=retained_pile)
        self.worker.start()
        self.listener_state.set("STARTING")
        self.listener_badge.configure(style="Stopped.TLabel")
        self.status.set("Starting capture service...")
        self.detail.set("Checking local port 9100.")

    def report_callback_exception(
        self,
        exception_type: type[BaseException],
        exception: BaseException,
        traceback_object: object,
    ) -> None:
        LOGGER.error(
            "Unhandled Tk callback exception",
            exc_info=(exception_type, exception, traceback_object),
        )
        try:
            self.status.set("Interface recovered from an error")
            self.detail.set(f"{exception_type.__name__}: {exception}")
        except tk.TclError:
            pass

    def stop(self) -> None:
        if self.worker is not None:
            self.worker.stop()
        self.listener_state.set("STOPPING")
        self.listener_badge.configure(style="Stopped.TLabel")
        self.status.set("Stopping...")

    def close(self) -> None:
        if self.worker is not None:
            self.worker.stop()
        self.root.destroy()

    def restore_main_window(self) -> None:
        """Make an incoming-job prompt reachable even if the app was minimized."""
        try:
            if self.root.state() in {"iconic", "withdrawn"}:
                self.root.deiconify()
            self.root.update_idletasks()
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after(900, lambda: self.root.attributes("-topmost", False))
        except tk.TclError:
            pass

    def action_dialog(
        self,
        title: str,
        heading: str,
        message: str,
        options: list[tuple[str, str, str]],
    ) -> str:
        self.restore_main_window()
        result = tk.StringVar(value="cancel")
        dialog = tk.Toplevel(self.root)
        dialog.withdraw()
        dialog.title(title)
        dialog.configure(background="#F2F7F3")
        dialog.resizable(False, False)

        card = ttk.Frame(dialog, style="Card.TFrame", padding=22)
        card.pack(fill="both", expand=True, padx=14, pady=14)
        ttk.Label(card, text=heading, style="Status.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text=message,
            style="Detail.TLabel",
            wraplength=520,
            justify="left",
        ).pack(anchor="w", pady=(7, 18))

        def choose(value: str) -> None:
            result.set(value)
            dialog.destroy()

        for value, label, style_name in options:
            ttk.Button(
                card,
                text=label,
                command=lambda selected=value: choose(selected),
                style=style_name,
            ).pack(fill="x", pady=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", lambda: choose("cancel"))
        dialog.update_idletasks()
        dialog_width = dialog.winfo_reqwidth()
        dialog_height = dialog.winfo_reqheight()
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dialog_width) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dialog_height) // 2
        x = min(max(20, x), max(20, screen_width - dialog_width - 20))
        y = min(max(20, y), max(20, screen_height - dialog_height - 60))
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        dialog.deiconify()
        dialog.lift()
        dialog.attributes("-topmost", True)
        dialog.wait_visibility()
        dialog.grab_set()
        dialog.focus_force()

        def release_topmost() -> None:
            try:
                dialog.attributes("-topmost", False)
            except tk.TclError:
                pass

        dialog.after(1200, release_topmost)
        dialog.bind("<Escape>", lambda _event: choose("cancel"))
        dialog.bell()
        self.root.wait_window(dialog)
        return result.get()

    def show_print_summary(
        self,
        summary: list[dict[str, object]],
        order_label_count: int,
        context: str,
    ) -> None:
        if not summary:
            return
        self.restore_main_window()
        dialog = tk.Toplevel(self.root)
        dialog.withdraw()
        dialog.title("Pile summary" if context == "pile" else "Sorted batch summary")
        dialog.configure(background="#F2F7F3")
        dialog.minsize(420, 320)

        card = ttk.Frame(dialog, style="Card.TFrame", padding=20)
        card.pack(fill="both", expand=True, padx=14, pady=14)
        ttk.Label(
            card,
            text="PRINTED LABEL SUMMARY",
            style="Eyebrow.TLabel",
        ).pack(anchor="w")
        ttk.Label(
            card,
            text=f"{order_label_count} order labels submitted",
            style="Status.TLabel",
        ).pack(anchor="w", pady=(7, 3))
        ttk.Label(
            card,
            text="Counts below exclude the inserted category-divider labels.",
            style="Detail.TLabel",
        ).pack(anchor="w", pady=(0, 12))

        text_frame = ttk.Frame(card, style="Card.TFrame")
        text_frame.pack(fill="both", expand=True)
        summary_text = tk.Text(
            text_frame,
            width=38,
            height=min(26, max(10, sum(len(group["sizes"]) + 2 for group in summary))),
            background="#F8FCF9",
            foreground="#294634",
            relief="flat",
            borderwidth=0,
            padx=16,
            pady=12,
            font=("Segoe UI", 11),
            cursor="arrow",
        )
        scrollbar = ttk.Scrollbar(
            text_frame,
            orient="vertical",
            command=summary_text.yview,
        )
        summary_text.configure(yscrollcommand=scrollbar.set)
        summary_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        summary_text.tag_configure(
            "heading",
            foreground="#166534",
            font=("Segoe UI", 12, "bold"),
            spacing1=6,
            spacing3=3,
        )
        summary_text.tag_configure(
            "size",
            lmargin1=22,
            lmargin2=22,
            spacing1=2,
            spacing3=2,
        )
        for group in summary:
            total = int(group["total"])
            summary_text.insert(
                "end",
                f"{group['heading']}  ({total})\n",
                "heading",
            )
            sizes = group["sizes"]
            if sizes:
                for size_entry in sizes:
                    count = int(size_entry["count"])
                    summary_text.insert(
                        "end",
                        f"{int(size_entry['size'])}  ────→  {count}\n",
                        "size",
                    )
            else:
                summary_text.insert("end", f"Unclassified  ────→  {total}\n", "size")
            summary_text.insert("end", "\n")
        summary_text.configure(state="disabled")

        ttk.Button(
            card,
            text="Close summary",
            command=dialog.destroy,
            style="Primary.TButton",
        ).pack(fill="x", pady=(14, 0))
        dialog.bind("<Escape>", lambda _event: dialog.destroy())

        dialog.update_idletasks()
        dialog_width = max(460, dialog.winfo_reqwidth())
        dialog_height = min(
            dialog.winfo_screenheight() - 100,
            max(420, dialog.winfo_reqheight()),
        )
        x = min(
            max(20, self.root.winfo_rootx() + 70),
            max(20, dialog.winfo_screenwidth() - dialog_width - 20),
        )
        y = min(
            max(20, self.root.winfo_rooty() + 45),
            max(20, dialog.winfo_screenheight() - dialog_height - 60),
        )
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        dialog.deiconify()
        dialog.lift()
        dialog.attributes("-topmost", True)
        dialog.focus_force()

        def release_topmost() -> None:
            try:
                dialog.attributes("-topmost", False)
            except tk.TclError:
                pass

        dialog.after(1200, release_topmost)

    def print_pile(self) -> None:
        if self.worker is None or not self.worker.is_alive():
            messagebox.showwarning(
                "Listener not running",
                "Start the listener before printing the pile.",
            )
            return
        if self.pile_label_count < 1:
            messagebox.showinfo("Pile is empty", "There are no batches in the pile.")
            return

        choice = self.action_dialog(
            "Print pile",
            "Print the complete pile?",
            (
                f"{self.pile_job_count} jobs containing "
                f"{self.pile_label_count} order labels are ready.\n\n"
                "Sorted printing combines and sorts every label globally. "
                "Normal printing preserves Job A, then Job B, in original order."
            ),
            [
                ("sorted", "Sort & Print Pile", "Primary.TButton"),
                ("original", "Print Pile Normally", "Secondary.TButton"),
                ("cancel", "Cancel", "Secondary.TButton"),
            ],
        )
        if choice in {"sorted", "original"}:
            self.worker.request_print_pile(choice)
            self.status.set("Preparing pile")
            self.detail.set("Combining all piled jobs into one print batch.")
            self.print_pile_button.configure(state="disabled")

    def poll_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                try:
                    self.handle_event(event)
                except Exception as error:
                    LOGGER.exception("Failed to handle UI event %r", event)
                    if (
                        event.get("type") == "decision_required"
                        and self.worker is not None
                    ):
                        self.worker.submit_decision("cancel")
                    self.status.set("Interface recovered from an error")
                    self.detail.set(f"{type(error).__name__}: {error}")
        except queue.Empty:
            pass
        finally:
            try:
                self.root.after(100, self.poll_events)
            except tk.TclError:
                pass

    def handle_event(self, event: dict[str, Any]) -> None:
        event_type = event["type"]
        if event_type == "listening":
            self.listener_state.set("LISTENING")
            self.listener_badge.configure(style="Listening.TLabel")
            self.status.set("Waiting for a Flipkart print batch")
            self.detail.set("Choose PDFCreator in Flipkart/QZ and fire the batch.")
            self.start_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
        elif event_type == "capturing":
            self.status.set("Batch received")
            self.detail.set(f"Capturing {event['byte_count']:,} bytes…")
        elif event_type == "captured":
            self.status.set(f"Detected {event['label_count']} ZPL labels")
            self.detail.set(
                "Label details extracted (SKU, product, quantity, payment, "
                "service, carrier, packaging, seller and AWB). Preparing batch."
            )
        elif event_type == "decision_required":
            choice = self.action_dialog(
                "New Flipkart batch",
                f"{event['label_count']} order labels received",
                (
                    "Choose how this batch should be handled. Adding it to the "
                    "pile saves it without printing."
                ),
                [
                    ("sorted", "Sort & Print", "Primary.TButton"),
                    ("original", "Print Original Order", "Secondary.TButton"),
                    ("pile", "Add to Pile", "Secondary.TButton"),
                    ("cancel", "Cancel — Do Not Print", "Secondary.TButton"),
                ],
            )
            if self.worker is not None:
                self.worker.submit_decision(choice)
            self.status.set("Processing your selection")
            self.detail.set("The captured batch remains saved locally.")
        elif event_type == "sorting":
            self.status.set("Sorting stage")
            self.detail.set(
                "Sorting by color, then BAGGY/PLAIN, then size; MIX goes last."
            )
        elif event_type == "printing":
            self.status.set(f"Printing {event['label_count']} labels…")
            self.detail.set(
                f"{event['order_label_count']} order labels + "
                f"{event['separator_count']} category labels; sending RAW ZPL "
                f"to {PHYSICAL_PRINTER}."
            )
        elif event_type == "printed":
            context = str(event.get("context", "batch"))
            self.status.set(
                "Pile sent successfully"
                if context == "pile"
                else "Batch sent successfully"
            )
            self.detail.set("Waiting for the next Flipkart print batch.")
            self.history.insert(
                "",
                0,
                values=(
                    datetime.now().strftime("%H:%M:%S"),
                    event["label_count"],
                    event["windows_job_id"],
                    (
                        "Pile submitted to TTP-244 Pro"
                        if context == "pile"
                        else "Submitted to TTP-244 Pro"
                    ),
                ),
                tags=("success",),
            )
            summary = event.get("summary")
            if isinstance(summary, list):
                self.show_print_summary(
                    summary,
                    int(event["order_label_count"]),
                    context,
                )
        elif event_type == "pile_updated":
            self.pile_job_count = int(event["job_count"])
            self.pile_label_count = int(event["label_count"])
            self.pile_state.set(
                f"Pile: {self.pile_job_count} jobs / "
                f"{self.pile_label_count} labels"
            )
            self.print_pile_button.configure(
                state="normal" if self.pile_label_count else "disabled"
            )
        elif event_type == "added_to_pile":
            self.status.set("Batch added to pile")
            self.detail.set(
                f"Stored {event['label_count']} labels without printing. "
                "Waiting for the next batch."
            )
            self.history.insert(
                "",
                0,
                values=(
                    datetime.now().strftime("%H:%M:%S"),
                    event["label_count"],
                    "—",
                    "Added to pile",
                ),
            )
        elif event_type == "batch_cancelled":
            self.status.set("Batch cancelled")
            self.detail.set(
                "The batch was saved locally but was not printed or added to the pile."
            )
        elif event_type == "pile_empty":
            self.status.set("Pile is empty")
            self.detail.set("Add one or more incoming batches before printing the pile.")
        elif event_type == "batch_error":
            context = str(event.get("context", "batch"))
            if context == "pile" and self.pile_label_count:
                self.print_pile_button.configure(state="normal")
                self.status.set("Pile print failed — pile retained")
            else:
                self.status.set("Batch failed")
            self.detail.set(event["message"])
            self.history.insert(
                "",
                0,
                values=(datetime.now().strftime("%H:%M:%S"), "—", "—", event["message"]),
                tags=("error",),
            )
            messagebox.showerror("Batch printing failed", event["message"])
        elif event_type == "fatal":
            self.listener_state.set("FAILED")
            self.listener_badge.configure(style="Error.TLabel")
            self.status.set("Capture service could not start")
            self.detail.set(event["message"])
            messagebox.showerror("Capture service failed", event["message"])
        elif event_type == "stopped":
            if self.listener_state.get() != "FAILED":
                self.listener_state.set("STOPPED")
                self.listener_badge.configure(style="Stopped.TLabel")
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            if self.status.get() == "Stopping...":
                self.status.set("Capture service stopped")


def main() -> None:
    configure_runtime_logging()
    LOGGER.info("Starting Flipkart Batch Printer")
    enable_high_dpi_rendering()
    root = tk.Tk()
    detected_dpi = root.winfo_fpixels("1i")
    root.tk.call("tk", "scaling", detected_dpi / 72.0)
    BatchPrinterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
