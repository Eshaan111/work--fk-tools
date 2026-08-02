from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import time
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


APPLICATION_DIRECTORY = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = APPLICATION_DIRECTORY / "print-middleware.json"


@dataclass(frozen=True)
class MiddlewareConfig:
    incoming_directory: Path
    processing_directory: Path
    archive_directory: Path
    failed_directory: Path
    poll_seconds: float
    stable_seconds: float
    pdfcreator_capture_enabled: bool
    pdfcreator_profile: str
    ocr_enabled: bool
    ocr_language: str
    ocr_dpi: int
    tesseract_command: str | None
    print_enabled: bool
    printer_name: str
    sumatra_path: Path | None
    print_settings: str


def resolve_path(value: str) -> Path:
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = APPLICATION_DIRECTORY / path
    return path.resolve()


def load_config(path: Path) -> MiddlewareConfig:
    with path.open("r", encoding="utf-8") as config_file:
        raw = json.load(config_file)

    directories = raw["directories"]
    ocr = raw.get("ocr", {})
    printing = raw.get("printing", {})
    capture = raw.get("pdfcreator_capture", {})
    sumatra_value = str(printing.get("sumatra_path", "")).strip()

    return MiddlewareConfig(
        incoming_directory=resolve_path(directories["incoming"]),
        processing_directory=resolve_path(directories["processing"]),
        archive_directory=resolve_path(directories["archive"]),
        failed_directory=resolve_path(directories["failed"]),
        poll_seconds=float(raw.get("poll_seconds", 1.0)),
        stable_seconds=float(raw.get("stable_seconds", 2.0)),
        pdfcreator_capture_enabled=bool(capture.get("enabled", True)),
        pdfcreator_profile=str(capture.get("profile", "DefaultGuid")).strip(),
        ocr_enabled=bool(ocr.get("enabled", True)),
        ocr_language=str(ocr.get("language", "eng")),
        ocr_dpi=int(ocr.get("dpi", 300)),
        tesseract_command=str(ocr.get("tesseract_command", "")).strip() or None,
        print_enabled=bool(printing.get("enabled", False)),
        printer_name=str(printing.get("printer_name", "")).strip(),
        sumatra_path=resolve_path(sumatra_value) if sumatra_value else None,
        print_settings=str(
            printing.get(
                "settings",
                "paper=auto,noscale,monochrome,disable-auto-rotation",
            )
        ).strip(),
    )


def ensure_directories(config: MiddlewareConfig) -> None:
    for directory in (
        config.incoming_directory,
        config.processing_directory,
        config.archive_directory,
        config.failed_directory,
    ):
        directory.mkdir(parents=True, exist_ok=True)


class PdfCreatorCapture:
    def __init__(self, config: MiddlewareConfig) -> None:
        self.config = config

    def __enter__(self) -> "PdfCreatorCapture":
        return self

    def __exit__(self, *_args: object) -> None:
        return None

    def capture_next(self, wait_seconds: int = 1) -> Path | None:
        if not self.config.pdfcreator_capture_enabled:
            return None

        job_name = (
            f"flipkart-{datetime.now():%Y%m%d-%H%M%S}-"
            f"{uuid.uuid4().hex[:8]}.pdf"
        )
        destination = self.config.incoming_directory / job_name
        capture_script = APPLICATION_DIRECTORY / "pdfcreator_capture.js"
        result = subprocess.run(
            [
                "cscript.exe",
                "//nologo",
                str(capture_script),
                str(destination),
                str(wait_seconds),
                self.config.pdfcreator_profile,
            ],
            check=False,
            capture_output=True,
            text=True,
            timeout=wait_seconds + 90,
        )
        output = result.stdout.strip()
        if result.returncode == 10:
            return None
        if result.returncode:
            detail = result.stderr.strip() or output or "No error details"
            raise RuntimeError(f"PDFCreator capture failed: {detail}")
        if not destination.is_file() or destination.stat().st_size == 0:
            raise RuntimeError(
                f"PDFCreator reported success but did not create: {destination}"
            )

        print(f"[captured] {destination.name}", flush=True)
        return destination


def utc_timestamp() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def unique_destination(directory: Path, name: str) -> Path:
    candidate = directory / name
    if not candidate.exists():
        return candidate

    stem = Path(name).stem
    suffix = Path(name).suffix
    counter = 2
    while True:
        candidate = directory / f"{stem}-{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def claim_pdf(source: Path, config: MiddlewareConfig) -> Path:
    destination = unique_destination(config.processing_directory, source.name)
    source.replace(destination)
    return destination


def extract_pdf(path: Path, config: MiddlewareConfig) -> dict[str, Any]:
    try:
        import fitz
    except ImportError as error:
        raise RuntimeError(
            "PyMuPDF is not installed. Run: python -m pip install -r requirements.txt"
        ) from error

    if config.ocr_enabled:
        try:
            import pytesseract
            from PIL import Image
        except ImportError as error:
            raise RuntimeError(
                "OCR packages are not installed. Run: "
                "python -m pip install -r requirements.txt"
            ) from error

        if config.tesseract_command:
            pytesseract.pytesseract.tesseract_cmd = config.tesseract_command

    document = fitz.open(path)
    pages: list[dict[str, Any]] = []
    scale = config.ocr_dpi / 72
    matrix = fitz.Matrix(scale, scale)

    try:
        for page_number, page in enumerate(document, start=1):
            embedded_text = page.get_text("text").strip()
            page_result: dict[str, Any] = {
                "page": page_number,
                "width_points": round(page.rect.width, 3),
                "height_points": round(page.rect.height, 3),
                "embedded_text": embedded_text,
            }

            if config.ocr_enabled:
                pixmap = page.get_pixmap(matrix=matrix, alpha=False)
                image = Image.frombytes(
                    "RGB",
                    (pixmap.width, pixmap.height),
                    pixmap.samples,
                )
                page_result["ocr_text"] = pytesseract.image_to_string(
                    image,
                    lang=config.ocr_language,
                    config="--psm 6",
                ).strip()

            pages.append(page_result)
    finally:
        document.close()

    return {
        "source_file": path.name,
        "sha256": sha256_file(path),
        "processed_at": utc_timestamp(),
        "page_count": len(pages),
        "pages": pages,
    }


def print_pdf(path: Path, config: MiddlewareConfig) -> None:
    if not config.print_enabled:
        return
    if not config.printer_name:
        raise RuntimeError("Printing is enabled but printer_name is empty.")
    if config.sumatra_path is None or not config.sumatra_path.is_file():
        raise RuntimeError(
            f"SumatraPDF was not found at: {config.sumatra_path or '<not configured>'}"
        )

    command = [
        str(config.sumatra_path),
        "-silent",
        "-print-to",
        config.printer_name,
    ]
    if config.print_settings:
        command.extend(["-print-settings", config.print_settings])
    command.append(str(path))

    result = subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode:
        detail = result.stderr.strip() or result.stdout.strip() or "No error details"
        raise RuntimeError(
            f"SumatraPDF printing failed with exit code {result.returncode}: {detail}"
        )


def process_pdf(source: Path, config: MiddlewareConfig) -> None:
    claimed = claim_pdf(source, config)
    sidecar = claimed.with_suffix(".ocr.json")

    try:
        extraction = extract_pdf(claimed, config)
        sidecar.write_text(
            json.dumps(extraction, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        print_pdf(claimed, config)

        archived_pdf = unique_destination(config.archive_directory, claimed.name)
        archived_sidecar = archived_pdf.with_suffix(".ocr.json")
        claimed.replace(archived_pdf)
        sidecar.replace(archived_sidecar)
        print(f"[processed] {archived_pdf.name}", flush=True)
    except Exception as error:
        failed_pdf = unique_destination(config.failed_directory, claimed.name)
        if claimed.exists():
            claimed.replace(failed_pdf)
        if sidecar.exists():
            failed_sidecar = failed_pdf.with_suffix(".ocr.json")
            sidecar.replace(failed_sidecar)
        error_path = failed_pdf.with_suffix(".error.txt")
        error_path.write_text(
            f"{utc_timestamp()}\n{type(error).__name__}: {error}\n",
            encoding="utf-8",
        )
        print(f"[failed] {source.name}: {error}", file=sys.stderr, flush=True)


def find_ready_pdfs(
    directory: Path,
    observations: dict[Path, tuple[int, float, float]],
    stable_seconds: float,
) -> list[Path]:
    now = time.monotonic()
    ready: list[Path] = []
    current_paths = set(directory.glob("*.pdf"))

    for stale_path in set(observations) - current_paths:
        observations.pop(stale_path, None)

    for path in sorted(current_paths):
        try:
            size = path.stat().st_size
        except FileNotFoundError:
            continue

        previous = observations.get(path)
        if previous is None or previous[0] != size:
            observations[path] = (size, now, path.stat().st_mtime)
            continue

        unchanged_since = previous[1]
        if size > 0 and now - unchanged_since >= stable_seconds:
            ready.append(path)
            observations.pop(path, None)

    return ready


def run(config: MiddlewareConfig, once: bool) -> None:
    ensure_directories(config)
    observations: dict[Path, tuple[int, float, float]] = {}

    print(f"Watching: {config.incoming_directory}", flush=True)
    print(
        f"PDFCreator capture: "
        f"{'enabled' if config.pdfcreator_capture_enabled else 'disabled'}; "
        f"OCR: {'enabled' if config.ocr_enabled else 'disabled'}; "
        f"printing: {'enabled' if config.print_enabled else 'disabled'}",
        flush=True,
    )

    with PdfCreatorCapture(config) as capture:
        while True:
            if not once:
                capture.capture_next(max(1, round(config.poll_seconds)))

            ready = find_ready_pdfs(
                config.incoming_directory,
                observations,
                config.stable_seconds,
            )
            for path in ready:
                process_pdf(path, config)

            if once:
                if not ready and any(config.incoming_directory.glob("*.pdf")):
                    time.sleep(config.stable_seconds)
                    continue
                return

            if not config.pdfcreator_capture_enabled:
                time.sleep(config.poll_seconds)


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Capture PDF print jobs, OCR them, and forward them to a printer."
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=DEFAULT_CONFIG_PATH,
        help=f"Configuration file (default: {DEFAULT_CONFIG_PATH.name})",
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="Process the current inbox and exit.",
    )
    return parser.parse_args()


def main() -> int:
    arguments = parse_arguments()
    config_path = arguments.config.resolve()
    if not config_path.exists():
        print(
            f"Configuration file not found: {config_path}\n"
            "Copy print-middleware.example.json to print-middleware.json first.",
            file=sys.stderr,
        )
        return 2

    try:
        run(load_config(config_path), arguments.once)
    except KeyboardInterrupt:
        print("\nStopped.", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
