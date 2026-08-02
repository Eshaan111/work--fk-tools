from __future__ import annotations

import argparse
import csv
import json
import re
import subprocess
from dataclasses import dataclass, asdict
from pathlib import Path


APPLICATION_DIRECTORY = Path(__file__).resolve().parent
RENDER_SCRIPT = APPLICATION_DIRECTORY / "zpl-preview" / "render-zpl.mjs"
ORDER_ID_PATTERN = re.compile(r"\bOD\d{16,22}\b", re.IGNORECASE)
ALLOWED_SIZES = (26, 28, 30, 32, 34, 36)


@dataclass(frozen=True)
class ExtractedLabel:
    batch_id: str
    label_number: int
    order_id: str | None
    sku: str | None
    product: str | None
    quantity: int | None
    payment_type: str | None
    service_type: str | None
    carrier: str | None
    packaging_instruction: str | None
    seller: str | None
    printed_at: str | None
    awb: str | None
    fit: str
    color: str | None
    size: int | None
    sort_tag: str
    is_mix: bool
    zpl_file: str
    preview_file: str | None


def split_printable_labels(zpl: str) -> list[str]:
    candidates = re.findall(r"\^XA.*?\^XZ", zpl, flags=re.IGNORECASE | re.DOTALL)
    return [
        candidate
        for candidate in candidates
        if re.search(
            r"\^(?:FD|GFA|BC|B[CQXR]|BX)",
            candidate,
            flags=re.IGNORECASE,
        )
    ]


def field_data(label: str) -> list[str]:
    fields = re.findall(
        r"\^FD(.*?)\^FS",
        label,
        flags=re.IGNORECASE | re.DOTALL,
    )
    return [
        re.sub(r"\s+", " ", field.replace(r"\&", " ")).strip()
        for field in fields
    ]


def first_match(pattern: re.Pattern[str], fields: list[str]) -> str | None:
    for field in fields:
        match = pattern.search(field)
        if match:
            return match.group(0).upper()
    return None


def extract_product(fields: list[str]) -> tuple[str | None, str | None, int | None]:
    for field in fields:
        if "|" not in field:
            continue
        candidate, product = (part.strip() for part in field.split("|", 1))
        if not (
            candidate
            and product
            and re.fullmatch(r"[A-Z0-9][A-Z0-9._/-]{3,}", candidate, re.I)
        ):
            continue

        quantity: int | None = None
        product_index = fields.index(field)
        for following_field in fields[product_index + 1 :]:
            if re.fullmatch(r"\d+", following_field):
                quantity = int(following_field)
                break
        return candidate, product, quantity
    return None, None, None


def extract_prefixed_value(fields: list[str], prefix: str) -> str | None:
    pattern = re.compile(rf"^{re.escape(prefix)}\s*:?[.]?\s*(.+)$", re.I)
    for field in fields:
        match = pattern.match(field)
        if match:
            return match.group(1).strip()
    return None


def extract_first_matching_field(
    fields: list[str], pattern: str
) -> str | None:
    matcher = re.compile(pattern, re.I)
    for field in fields:
        if matcher.fullmatch(field):
            return field.strip()
    return None


def extract_seller(fields: list[str]) -> str | None:
    seller_details = extract_prefixed_value(fields, "Sold By")
    if not seller_details:
        return None
    return seller_details.split(",", 1)[0].strip()


def extract_awb(fields: list[str]) -> str | None:
    for field in fields:
        match = re.search(r"\bAWB\s+No\.?\s*([A-Z0-9-]{8,})", field, re.I)
        if match:
            return match.group(1).upper()

    counts: dict[str, int] = {}
    for field in fields:
        if re.fullmatch(r"[A-Z]{1,4}\d{8,}[A-Z]?", field, re.I):
            normalized = field.upper()
            counts[normalized] = counts.get(normalized, 0) + 1
    if counts:
        return max(counts, key=counts.get)
    return None


def classify_for_sorting(
    sku: str | None,
    product: str | None,
) -> tuple[str, str | None, int | None, str, bool]:
    """Classify a label without changing any of its printable ZPL."""
    normalized_sku = (sku or "").upper()
    normalized_product = (product or "").upper()

    fit = (
        "BAGGY"
        if "BAGGY" in normalized_sku or "RELAXED" in normalized_product
        else "PLAIN"
    )

    color: str | None = None
    for candidate, terms in (
        ("ICE", ("ICE", "BLUE")),
        ("BEIGE", ("CREAM", "BEIGE")),
        ("BLACK", ("BLACK",)),
        ("WHITE", ("WHITE",)),
    ):
        if any(term in normalized_sku for term in terms):
            color = candidate
            break

    # ICE and BEIGE are always treated as baggy, per the business rule.
    if color in {"ICE", "BEIGE"}:
        fit = "BAGGY"

    size_tokens = re.findall(
        r"(?:^|[-_])(26|28|30|32|34|36)(?=$|[-_])",
        normalized_sku,
    )
    size: int | None = None
    if len(size_tokens) == 1:
        size = 32 if re.search(r"_39$", normalized_sku) else int(size_tokens[0])
    elif not size_tokens and re.search(r"_39$", normalized_sku):
        size = 32
    # Two or more recognized size tokens are deliberately left unresolved,
    # which sends the label to MIX even when the tokens repeat the same size.

    is_mix = color is None or size is None
    sort_tag = "MIX" if is_mix else f"{color}-{fit}-{size}"
    return fit, color, size, sort_tag, is_mix


def render_label(zpl_path: Path, preview_path: Path) -> None:
    result = subprocess.run(
        [
            "node",
            str(RENDER_SCRIPT),
            str(zpl_path),
            str(preview_path),
        ],
        check=False,
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode:
        detail = result.stderr.strip() or result.stdout.strip() or "No error details"
        raise RuntimeError(f"ZPL preview rendering failed: {detail}")


def process_zpl_job(
    source: Path,
    output_directory: Path,
    render_previews: bool = True,
) -> list[ExtractedLabel]:
    zpl = source.read_text(encoding="utf-8", errors="replace")
    labels = split_printable_labels(zpl)
    if not labels:
        raise ValueError(f"No printable ^XA...^XZ labels found in {source.name}.")

    output_directory.mkdir(parents=True, exist_ok=True)
    extracted: list[ExtractedLabel] = []
    batch_id = output_directory.name

    for index, label in enumerate(labels, start=1):
        fields = field_data(label)
        order_id = first_match(ORDER_ID_PATTERN, fields)
        sku, product, quantity = extract_product(fields)
        awb = extract_awb(fields)
        payment_type = extract_first_matching_field(
            fields, r"(?:COD|PREPAID|POSTPAID)"
        )
        service_type = extract_first_matching_field(
            fields, r"(?:STD|STANDARD|EXPRESS|PRIORITY)"
        )
        carrier = extract_first_matching_field(fields, r".+\bLogistics\b")
        packaging_instruction = extract_first_matching_field(
            fields, r"Use\s+.+\s+Packaging"
        )
        seller = extract_seller(fields)
        printed_at = extract_prefixed_value(fields, "Printed at")
        fit, color, size, sort_tag, is_mix = classify_for_sorting(sku, product)
        identifier = order_id or f"label-{index:03d}"
        safe_identifier = re.sub(r"[^A-Za-z0-9._-]", "_", identifier)

        zpl_path = output_directory / f"{index:03d}-{safe_identifier}.zpl"
        preview_path = output_directory / f"{index:03d}-{safe_identifier}.png"
        zpl_path.write_text(label, encoding="utf-8")

        rendered_path: str | None = None
        if render_previews:
            render_label(zpl_path, preview_path)
            rendered_path = str(preview_path)

        record = ExtractedLabel(
            batch_id=batch_id,
            label_number=index,
            order_id=order_id,
            sku=sku,
            product=product,
            quantity=quantity,
            payment_type=payment_type,
            service_type=service_type,
            carrier=carrier,
            packaging_instruction=packaging_instruction,
            seller=seller,
            printed_at=printed_at,
            awb=awb,
            fit=fit,
            color=color,
            size=size,
            sort_tag=sort_tag,
            is_mix=is_mix,
            zpl_file=str(zpl_path),
            preview_file=rendered_path,
        )
        extracted.append(record)
        (output_directory / f"{index:03d}-{safe_identifier}.json").write_text(
            json.dumps(asdict(record), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    summary_path = output_directory / "labels.json"
    summary_path.write_text(
        json.dumps([asdict(record) for record in extracted], indent=2),
        encoding="utf-8",
    )

    with (output_directory / "labels.csv").open(
        "w",
        encoding="utf-8-sig",
        newline="",
    ) as csv_file:
        writer = csv.DictWriter(
            csv_file,
            fieldnames=[
                "batch_id",
                "label_number",
                "order_id",
                "sku",
                "product",
                "quantity",
                "payment_type",
                "service_type",
                "carrier",
                "packaging_instruction",
                "seller",
                "printed_at",
                "awb",
                "fit",
                "color",
                "size",
                "sort_tag",
                "is_mix",
                "zpl_file",
                "preview_file",
            ],
        )
        writer.writeheader()
        writer.writerows(asdict(record) for record in extracted)

    return extracted


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Split, extract, and locally render labels from a ZPL print job."
    )
    parser.add_argument("source", type=Path)
    parser.add_argument("output_directory", type=Path)
    parser.add_argument("--no-render", action="store_true")
    return parser.parse_args()


def main() -> int:
    arguments = parse_arguments()
    records = process_zpl_job(
        arguments.source.resolve(),
        arguments.output_directory.resolve(),
        render_previews=not arguments.no_render,
    )
    print(json.dumps([asdict(record) for record in records], indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
