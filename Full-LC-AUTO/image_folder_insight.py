from __future__ import annotations

import re
from datetime import datetime
from collections import OrderedDict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_PATH = PROJECT_ROOT / "image_folder_insight.xlsx"
# LAPTOP_NAME = os.getenv("ASUS", "VAIO").upper()
LAPTOP_NAME = "ASUS"
# LAPTOP_NAME = "VAIO"

BRAND_CODE_MAP = OrderedDict(
    [
        ("STAR", "STARVIELLE"),
        ("GENZ", "GENZ VANE"),
        ("IND", "INDIVANE"),
        ("FADE", "FADEVIELLE"),
        ("FLEE", "FLEECRANE"),
    ]
)
PROFILE_BRAND_CODES = OrderedDict(
    [
        ("prabhu", ("STAR", "GENZ")),
        ("seema", ("FADE", "FLEE", "IND")),
    ]
)
SURFACE_FOLDER_SUFFIX = OrderedDict(
    [
        ("flipkart", "'f"),
        ("shopsy", "'s"),
    ]
)
FOLDER_SUFFIX_SURFACE = {
    suffix.upper(): surface for surface, suffix in SURFACE_FOLDER_SUFFIX.items()
}
KIND_DIRECTORIES_ASUS = OrderedDict(
    [
        ("Beige", Path(r"C:\work-mom\LISTING IMAGES AUTOMATED\BEIGE")),
        ("Ice", Path(r"C:\work-mom\LISTING IMAGES AUTOMATED\ICE")),
        ("Black-baggy", Path(r"C:\work-mom\LISTING IMAGES AUTOMATED\Black-baggy")),
        ("Black-Plain", Path(r"C:\work-mom\LISTING IMAGES AUTOMATED\Black-Plain")),
        ("White-Plain", Path(r"C:\work-mom\LISTING IMAGES AUTOMATED\White")),
        ("Trouser", Path(r"C:\work-mom\LISTING IMAGES AUTOMATED\Trouser")),
        ("Shorts", Path(r"C:\work-mom\LISTING IMAGES AUTOMATED\Shorts")),
    ]
)
KIND_DIRECTORIES_VAIO = OrderedDict(
    [
        ("Beige", Path(r"G:\Other computers\My Laptop\work-mom\LISTING IMAGES AUTOMATED\BEIGE")),
        ("Ice", Path(r"G:\Other computers\My Laptop\work-mom\LISTING IMAGES AUTOMATED\ICE")),
        ("Black-baggy", Path(r"G:\Other computers\My Laptop\work-mom\LISTING IMAGES AUTOMATED\Black-baggy")),
        ("Black-Plain", Path(r"G:\Other computers\My Laptop\work-mom\LISTING IMAGES AUTOMATED\Black-Plain")),
        ("White-Plain", Path(r"G:\Other computers\My Laptop\work-mom\LISTING IMAGES AUTOMATED\White")),
        ("Trouser", Path(r"G:\Other computers\My Laptop\work-mom\LISTING IMAGES AUTOMATED\Trouser")),
        ("Shorts", Path(r"G:\Other computers\My Laptop\work-mom\LISTING IMAGES AUTOMATED")),
    ]
)
LAPTOP_CONFIGS = {
    "ASUS": {
        "kind_directories": KIND_DIRECTORIES_ASUS,
    },
    "VAIO": {
        "kind_directories": KIND_DIRECTORIES_VAIO,
    },
}
if LAPTOP_NAME not in LAPTOP_CONFIGS:
    raise ValueError(f"Unknown LAPTOP_NAME '{LAPTOP_NAME}'. Choose ASUS or VAIO.")
ACTIVE_LAPTOP_CONFIG = LAPTOP_CONFIGS[LAPTOP_NAME]
KIND_DIRECTORIES = ACTIVE_LAPTOP_CONFIG["kind_directories"]
TOKEN_PATTERN = re.compile(r"([A-Z]+)('?[SF])?$")

THIN_BORDER = Border(
    left=Side(style="thin", color="D0D7DE"),
    right=Side(style="thin", color="D0D7DE"),
    top=Side(style="thin", color="D0D7DE"),
    bottom=Side(style="thin", color="D0D7DE"),
)
HEADER_FILL = PatternFill(fill_type="solid", fgColor="1F4E78")
KIND_FILL = PatternFill(fill_type="solid", fgColor="EAF3FF")
NONZERO_FILL = PatternFill(fill_type="solid", fgColor="D9EAD3")
ZERO_FILL = PatternFill(fill_type="solid", fgColor="F3F4F6")
NOTES_LABEL_FILL = PatternFill(fill_type="solid", fgColor="FCE5CD")
WARNING_FILL = PatternFill(fill_type="solid", fgColor="FDE9D9")
HEADER_FONT = Font(bold=True, color="FFFFFF")
LABEL_FONT = Font(bold=True, color="1F1F1F")
ZERO_FONT = Font(color="7A7A7A")
NONZERO_FONT = Font(bold=True, color="215E21")
CENTER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
LEFT_ALIGNMENT = Alignment(horizontal="left", vertical="center")

ACCOUNT_BRAND_COLUMNS: list[tuple[str, str, str]] = []
for account_name, brand_codes in PROFILE_BRAND_CODES.items():
    for brand_code in brand_codes:
        brand_name = BRAND_CODE_MAP[brand_code]
        for surface_name in SURFACE_FOLDER_SUFFIX:
            ACCOUNT_BRAND_COLUMNS.append((account_name, brand_name, surface_name))

BRAND_TO_ACCOUNT = {
    BRAND_CODE_MAP[brand_code]: account_name
    for account_name, brand_codes in PROFILE_BRAND_CODES.items()
    for brand_code in brand_codes
}


def parse_folder_number(folder_name: str) -> int:
    first_token = folder_name.split("-", maxsplit=1)[0].strip()
    if not first_token.isdigit():
        raise ValueError(f"Folder name must start with a number: {folder_name}")
    return int(first_token)


def parse_brand_folder_token(token: str, folder_name: str) -> list[tuple[str, str]]:
    normalized_token = token.strip().upper()
    token_match = TOKEN_PATTERN.fullmatch(normalized_token)
    if token_match is None:
        raise ValueError(f"Unknown brand code '{token}' in folder '{folder_name}'")

    brand_code, suffix = token_match.groups()
    if brand_code not in BRAND_CODE_MAP:
        raise ValueError(f"Unknown brand code '{token}' in folder '{folder_name}'")

    brand_name = BRAND_CODE_MAP[brand_code]
    if not suffix:
        return [(brand_name, surface_name) for surface_name in SURFACE_FOLDER_SUFFIX]

    normalized_suffix = suffix if suffix.startswith("'") else f"'{suffix}"
    surface_name = FOLDER_SUFFIX_SURFACE.get(normalized_suffix.upper())
    if surface_name is None:
        raise ValueError(f"Unknown brand surface suffix '{token}' in folder '{folder_name}'")

    return [(brand_name, surface_name)]


def parse_exhausted_brand_surfaces(folder_name: str) -> list[tuple[str, str]]:
    parse_folder_number(folder_name)
    exhausted_pairs: list[tuple[str, str]] = []
    tokens = [part.strip() for part in folder_name.split("-")[1:] if part.strip()]
    for token in tokens:
        exhausted_pairs.extend(parse_brand_folder_token(token, folder_name))
    return exhausted_pairs


def scan_kind_directory(
    kind_name: str,
    folder_root: Path,
) -> tuple[dict[tuple[str, str, str], int], list[str], int]:
    counts = {
        (account_name, brand_name, surface_name): 0
        for account_name, brand_name, surface_name in ACCOUNT_BRAND_COLUMNS
    }
    warnings: list[str] = []
    total_numbered_folders = 0

    if not folder_root.exists():
        warnings.append(f"{kind_name}: missing directory -> {folder_root}")
        return counts, warnings, total_numbered_folders
    if not folder_root.is_dir():
        warnings.append(f"{kind_name}: not a directory -> {folder_root}")
        return counts, warnings, total_numbered_folders

    for child in sorted(folder_root.iterdir(), key=lambda path: path.name.lower()):
        if not child.is_dir():
            continue

        try:
            parse_folder_number(child.name)
        except ValueError as exc:
            warnings.append(f"{kind_name}: skipped folder '{child.name}' ({exc})")
            continue

        total_numbered_folders += 1

        try:
            exhausted_pairs = parse_exhausted_brand_surfaces(child.name)
        except ValueError as exc:
            warnings.append(f"{kind_name}: skipped folder tokens in '{child.name}' ({exc})")
            continue

        for brand_name, surface_name in exhausted_pairs:
            account_name = BRAND_TO_ACCOUNT.get(brand_name)
            if account_name is None:
                warnings.append(
                    f"{kind_name}: skipped brand '{brand_name}' from folder '{child.name}' because no account mapping exists"
                )
                continue
            counts[(account_name, brand_name, surface_name)] += 1

    return counts, warnings, total_numbered_folders


def autosize_columns(worksheet) -> None:
    widths: dict[int, int] = {}
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            widths[cell.column] = max(widths.get(cell.column, 0), len(str(cell.value)))
    for column_index, width in widths.items():
        worksheet.column_dimensions[worksheet.cell(row=1, column=column_index).column_letter].width = min(width + 3, 32)


def style_header_cell(cell) -> None:
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.border = THIN_BORDER
    cell.alignment = CENTER_ALIGNMENT


def style_kind_cell(cell) -> None:
    cell.font = LABEL_FONT
    cell.fill = KIND_FILL
    cell.border = THIN_BORDER
    cell.alignment = LEFT_ALIGNMENT


def style_count_cell(cell, value: int) -> None:
    cell.border = THIN_BORDER
    cell.alignment = CENTER_ALIGNMENT
    if value > 0:
        cell.fill = NONZERO_FILL
        cell.font = NONZERO_FONT
    else:
        cell.fill = ZERO_FILL
        cell.font = ZERO_FONT


def build_workbook() -> Workbook:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Image Folder Insight"

    header_cell = worksheet.cell(row=1, column=1, value="Kind")
    style_header_cell(header_cell)
    for column_index, (account_name, brand_name, surface_name) in enumerate(ACCOUNT_BRAND_COLUMNS, start=2):
        header = f"{account_name.title()} - {brand_name} - {surface_name.title()}"
        cell = worksheet.cell(row=1, column=column_index, value=header)
        style_header_cell(cell)

    available_sheet = workbook.create_sheet("Available Options")
    available_kind_header = available_sheet.cell(row=1, column=1, value="Kind")
    style_header_cell(available_kind_header)
    total_folders_header = available_sheet.cell(row=1, column=2, value="Total Numbered Folders")
    style_header_cell(total_folders_header)
    for column_index, (account_name, brand_name, surface_name) in enumerate(ACCOUNT_BRAND_COLUMNS, start=3):
        header = f"{account_name.title()} - {brand_name} - {surface_name.title()}"
        cell = available_sheet.cell(row=1, column=column_index, value=header)
        style_header_cell(cell)

    warning_messages: list[str] = []
    for row_index, (kind_name, folder_root) in enumerate(KIND_DIRECTORIES.items(), start=2):
        kind_cell = worksheet.cell(row=row_index, column=1, value=kind_name)
        style_kind_cell(kind_cell)
        available_kind_cell = available_sheet.cell(row=row_index, column=1, value=kind_name)
        style_kind_cell(available_kind_cell)

        counts, warnings, total_numbered_folders = scan_kind_directory(kind_name, folder_root)
        warning_messages.extend(warnings)

        total_cell = available_sheet.cell(row=row_index, column=2, value=total_numbered_folders)
        style_count_cell(total_cell, total_numbered_folders)

        for column_index, key in enumerate(ACCOUNT_BRAND_COLUMNS, start=2):
            value = counts[key]
            cell = worksheet.cell(row=row_index, column=column_index, value=value)
            style_count_cell(cell, value)

        for column_index, key in enumerate(ACCOUNT_BRAND_COLUMNS, start=3):
            available_value = max(total_numbered_folders - counts[key], 0)
            cell = available_sheet.cell(row=row_index, column=column_index, value=available_value)
            style_count_cell(cell, available_value)

    worksheet.freeze_panes = "B2"
    worksheet.sheet_view.showGridLines = False
    worksheet.row_dimensions[1].height = 28
    autosize_columns(worksheet)

    available_sheet.freeze_panes = "C2"
    available_sheet.sheet_view.showGridLines = False
    available_sheet.row_dimensions[1].height = 28
    autosize_columns(available_sheet)

    notes = workbook.create_sheet("Notes")
    notes_rows = [
        ("Generated From", "Folder names under the configured image roots"),
        ("Rule", "Each exhausted brand/surface token in a numbered folder counts as one successful listing"),
        ("Suffix 'f", "Flipkart"),
        ("Suffix 's", "Shopsy"),
        ("No suffix", "Counts for both Flipkart and Shopsy"),
        ("Accounts", "Brand ownership is inferred from PROFILE_BRAND_CODES in main.py"),
        ("Available Options", "Each value is total numbered folders minus exhausted count for that account/brand/surface"),
    ]
    for row_index, (label, value) in enumerate(notes_rows, start=1):
        label_cell = notes.cell(row=row_index, column=1, value=label)
        label_cell.font = LABEL_FONT
        label_cell.fill = NOTES_LABEL_FILL
        label_cell.border = THIN_BORDER
        label_cell.alignment = LEFT_ALIGNMENT

        value_cell = notes.cell(row=row_index, column=2, value=value)
        value_cell.border = THIN_BORDER
        value_cell.alignment = LEFT_ALIGNMENT

    if warning_messages:
        warning_header = notes.cell(row=8, column=1, value="Warnings")
        warning_header.font = LABEL_FONT
        warning_header.fill = WARNING_FILL
        warning_header.border = THIN_BORDER
        warning_header.alignment = LEFT_ALIGNMENT
        for row_index, warning in enumerate(warning_messages, start=9):
            warning_cell = notes.cell(row=row_index, column=1, value=warning)
            warning_cell.border = THIN_BORDER
            warning_cell.alignment = LEFT_ALIGNMENT

    notes.sheet_view.showGridLines = False
    autosize_columns(notes)
    return workbook


def main() -> None:
    workbook = build_workbook()
    try:
        workbook.save(OUTPUT_PATH)
        print(f"Saved insight workbook: {OUTPUT_PATH}")
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_path = PROJECT_ROOT / f"image_folder_insight_{timestamp}.xlsx"
        workbook.save(fallback_path)
        print(
            "Primary workbook is locked. "
            f"Saved insight workbook to fallback path: {fallback_path}"
        )


if __name__ == "__main__":
    main()
