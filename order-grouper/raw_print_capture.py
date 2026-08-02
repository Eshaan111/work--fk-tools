from __future__ import annotations

import argparse
import hashlib
import json
import socket
from datetime import datetime, timezone
from pathlib import Path


APPLICATION_DIRECTORY = Path(__file__).resolve().parent
DEFAULT_RAW_DIRECTORY = APPLICATION_DIRECTORY / "print-jobs" / "raw"
DEFAULT_PDF_INBOX = APPLICATION_DIRECTORY / "print-jobs" / "incoming"
DEFAULT_LABEL_DIRECTORY = APPLICATION_DIRECTORY / "print-jobs" / "labels"


def utc_timestamp() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def detect_format(data: bytes) -> tuple[str, str]:
    stripped = data.lstrip(b"\x00\x09\x0a\x0c\x0d\x20")
    upper = stripped[:4096].upper()

    if stripped.startswith(b"%PDF-"):
        return "pdf", ".pdf"
    if b"^XA" in upper or b"^GFA" in upper or b"~DG" in upper:
        return "zpl", ".zpl"
    if (
        upper.startswith(b"SIZE ")
        or b"\r\nSIZE " in upper
        or (b"\r\nCLS" in upper and b"\r\nPRINT" in upper)
    ):
        return "tspl", ".tspl"
    if upper.startswith(b"N\r\n") and b"\r\nP" in upper:
        return "epl", ".epl"
    if stripped.startswith((b"\x1b@", b"\x1b\x40")):
        return "escpos", ".escpos"
    if stripped.startswith((b"\x1b%-12345X", b"@PJL")):
        return "pcl-or-postscript", ".prn"
    if stripped.startswith(b"%!PS"):
        return "postscript", ".ps"
    return "unknown", ".bin"


def save_job(
    data: bytes,
    raw_directory: Path,
    pdf_inbox: Path,
    labels_directory: Path,
    peer: tuple[str, int],
) -> tuple[Path, dict[str, object]]:
    job_format, suffix = detect_format(data)
    job_id = f"{datetime.now():%Y%m%d-%H%M%S-%f}"
    raw_directory.mkdir(parents=True, exist_ok=True)
    pdf_inbox.mkdir(parents=True, exist_ok=True)

    destination = raw_directory / f"raw-{job_id}{suffix}"
    destination.write_bytes(data)

    metadata: dict[str, object] = {
        "job_id": job_id,
        "received_at": utc_timestamp(),
        "peer": f"{peer[0]}:{peer[1]}",
        "format": job_format,
        "byte_count": len(data),
        "sha256": hashlib.sha256(data).hexdigest(),
        "starts_with_hex": data[:64].hex(" "),
        "raw_file": str(destination),
    }

    if job_format == "pdf":
        pdf_path = pdf_inbox / f"raw-{job_id}.pdf"
        pdf_path.write_bytes(data)
        metadata["pdf_inbox_file"] = str(pdf_path)
    elif job_format == "zpl":
        try:
            from zpl_job_processor import process_zpl_job

            output_directory = labels_directory / job_id
            labels = process_zpl_job(
                destination,
                output_directory,
                render_previews=False,
            )
            metadata["label_count"] = len(labels)
            metadata["labels_directory"] = str(output_directory)
            metadata["labels_summary"] = str(output_directory / "labels.json")
        except Exception as error:
            metadata["label_processing_error"] = f"{type(error).__name__}: {error}"

    metadata_path = destination.with_suffix(destination.suffix + ".json")
    metadata_path.write_text(
        json.dumps(metadata, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return destination, metadata


def receive_connection(
    connection: socket.socket,
    idle_timeout: float,
    maximum_bytes: int,
) -> bytes:
    connection.settimeout(idle_timeout)
    chunks: list[bytes] = []
    byte_count = 0

    while True:
        try:
            chunk = connection.recv(64 * 1024)
        except socket.timeout:
            break
        if not chunk:
            break

        byte_count += len(chunk)
        if byte_count > maximum_bytes:
            raise RuntimeError(
                f"Print job exceeded the configured {maximum_bytes} byte limit."
            )
        chunks.append(chunk)

    return b"".join(chunks)


def serve(
    host: str,
    port: int,
    raw_directory: Path,
    pdf_inbox: Path,
    labels_directory: Path,
    idle_timeout: float,
    maximum_bytes: int,
    once: bool,
) -> None:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as server:
        server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        server.bind((host, port))
        server.listen(8)
        print(f"Raw print capture listening on {host}:{port}", flush=True)

        while True:
            connection, peer = server.accept()
            with connection:
                try:
                    data = receive_connection(
                        connection,
                        idle_timeout=idle_timeout,
                        maximum_bytes=maximum_bytes,
                    )
                    if not data:
                        print(f"[empty connection] {peer[0]}:{peer[1]}", flush=True)
                    else:
                        destination, metadata = save_job(
                            data,
                            raw_directory=raw_directory,
                            pdf_inbox=pdf_inbox,
                            labels_directory=labels_directory,
                            peer=peer,
                        )
                        print(
                            f"[captured] {destination.name}: "
                            f"{metadata['format']}, {metadata['byte_count']} bytes",
                            flush=True,
                        )
                except Exception as error:
                    print(
                        f"[capture failed] {peer[0]}:{peer[1]}: {error}",
                        flush=True,
                    )

            if once:
                return


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Capture raw Windows printer jobs over a local TCP/IP port."
    )
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=9100)
    parser.add_argument("--raw-directory", type=Path, default=DEFAULT_RAW_DIRECTORY)
    parser.add_argument("--pdf-inbox", type=Path, default=DEFAULT_PDF_INBOX)
    parser.add_argument(
        "--labels-directory",
        type=Path,
        default=DEFAULT_LABEL_DIRECTORY,
    )
    parser.add_argument("--idle-timeout", type=float, default=2.0)
    parser.add_argument("--maximum-bytes", type=int, default=100 * 1024 * 1024)
    parser.add_argument("--once", action="store_true")
    return parser.parse_args()


def main() -> int:
    arguments = parse_arguments()
    serve(
        host=arguments.host,
        port=arguments.port,
        raw_directory=arguments.raw_directory.resolve(),
        pdf_inbox=arguments.pdf_inbox.resolve(),
        labels_directory=arguments.labels_directory.resolve(),
        idle_timeout=arguments.idle_timeout,
        maximum_bytes=arguments.maximum_bytes,
        once=arguments.once,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
