from __future__ import annotations

import argparse
from pathlib import Path

from zpl_job_processor import split_printable_labels


def print_raw_zpl(source: Path, printer_name: str) -> tuple[int, int]:
    try:
        import win32print
    except ImportError as error:
        raise RuntimeError(
            "pywin32 is not installed. Run: python -m pip install -r requirements.txt"
        ) from error

    data = source.read_bytes()
    if not data.strip():
        raise ValueError(f"ZPL file is empty: {source}")

    zpl = data.decode("utf-8", errors="replace")
    label_count = len(split_printable_labels(zpl))
    if label_count == 0:
        raise ValueError(f"No printable ZPL labels were found in: {source}")

    printer = win32print.OpenPrinter(printer_name)
    try:
        details = win32print.GetPrinter(printer, 2)
        port_name = str(details.get("pPortName", ""))
        if port_name == "FlipkartRawCapturePort":
            raise RuntimeError(
                "Refusing to print back into the Flipkart capture port. "
                "Choose the physical TTP-244 Pro queue."
            )

        job_id = win32print.StartDocPrinter(
            printer,
            1,
            (f"Flipkart ZPL Batch - {source.stem}", None, "RAW"),
        )
        try:
            win32print.StartPagePrinter(printer)
            written = win32print.WritePrinter(printer, data)
            win32print.EndPagePrinter(printer)
        finally:
            win32print.EndDocPrinter(printer)
    finally:
        win32print.ClosePrinter(printer)

    if written != len(data):
        raise RuntimeError(
            f"Only {written} of {len(data)} bytes were accepted by the spooler."
        )
    return job_id, label_count


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Send a captured ZPL batch unchanged to a physical printer."
    )
    parser.add_argument("source", type=Path)
    parser.add_argument(
        "--printer",
        default="TSC TTP-244 Pro",
        help="Physical Windows printer queue name.",
    )
    return parser.parse_args()


def main() -> int:
    arguments = parse_arguments()
    source = arguments.source.resolve()
    job_id, label_count = print_raw_zpl(source, arguments.printer)
    print(
        f"Submitted Windows job {job_id}: {label_count} labels from {source.name} "
        f"to {arguments.printer}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
