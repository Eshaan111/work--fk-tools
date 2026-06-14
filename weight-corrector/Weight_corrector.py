import csv
import tkinter as tk
from copy import deepcopy
from pathlib import Path
from tkinter import filedialog, messagebox
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


SOURCE_COLUMNS = [
    ("FSN", 8),
    ("SKU", 9),
    ("Product Title", 10),
    ("Package Length (cm)", 32),
    ("Package Breadth (cm)", 33),
    ("Package Height (cm)", 34),
    ("Package Weight (kg)", 35),
    ("Ready to Make", 36),
]

DEFAULT_TEMPLATE_NAME = "WEIGHT TEMPLATE.xlsx"
FLAG_COLUMN_NAME = "WRONG_WEIGHT_FLAG"
CURRENT_CONTEXT_NAME = "current_context.xlsx"
WRONG_WEIGHT_FILE_NAME = "WRONG_WEIGHT_FILE.xlsx"
XML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

TEMPLATE_KIND_MAP = {
    "WHITE": "WHITE",
    "ICE": "ICE",
    "BEIGE": "BEIGE",
    "BLACK-BAGGY": "BLACK BAGGY",
    "BLACK-PLAIN": "BLACK-PLAIN",
}


def classify_kind_of_jeans(sku: str, product_title: str) -> str:
    sku_lower = sku.lower()
    title_lower = product_title.lower()

    if "white" in sku_lower:
        return "WHITE"
    if "ice" in sku_lower or "blue" in sku_lower:
        return "ICE"
    if "beige" in sku_lower or "cream" in sku_lower:
        return "BEIGE"
    if "baggy" in sku_lower:
        return "BLACK-BAGGY"
    if "black" in sku_lower and "relaxed" in title_lower:
        return "BLACK-BAGGY"
    if "black" in sku_lower:
        return "BLACK-PLAIN"
    return "MIX"


def read_selected_columns(csv_path: Path) -> list[list[str]]:
    template_rows_by_kind = load_template_rows_by_kind(Path(__file__).with_name(DEFAULT_TEMPLATE_NAME))
    rows: list[list[str]] = []

    with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.reader(file)
        header = next(reader, None)
        has_leading_flag_column = bool(header and header[0] == FLAG_COLUMN_NAME)

        for row in reader:
            normalized_row = row[1:] if has_leading_flag_column and row else row
            selected_row = []
            for _, column_number in SOURCE_COLUMNS:
                index = column_number - 1
                selected_row.append(normalized_row[index] if index < len(normalized_row) else "")
            kind_of_jeans = classify_kind_of_jeans(selected_row[1], selected_row[2])
            wrong_weight_flag = get_wrong_weight_flag(
                kind_of_jeans,
                selected_row[6],
                template_rows_by_kind,
            )
            selected_row.append(kind_of_jeans)
            selected_row.append(wrong_weight_flag)
            rows.append(selected_row)

    return rows


def load_template_workbook_details(template_path: Path) -> tuple[dict[str, str], str, ET.Element, ET.Element]:
    namespace = {"main": XML_NS}

    with ZipFile(template_path) as workbook_zip:
        workbook_xml = ET.fromstring(workbook_zip.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(workbook_zip.read("xl/_rels/workbook.xml.rels"))

        rel_map = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels_xml
            if rel.tag.endswith("Relationship")
        }

        first_sheet = workbook_xml.find("main:sheets/main:sheet", namespace)
        if first_sheet is None:
            raise ValueError("No worksheets found in template workbook.")

        relationship_id = first_sheet.attrib[
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        ]
        sheet_path = f"xl/{rel_map[relationship_id]}"
        sheet_xml = ET.fromstring(workbook_zip.read(sheet_path))

        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in workbook_zip.namelist():
            shared_xml = ET.fromstring(workbook_zip.read("xl/sharedStrings.xml"))
            for item in shared_xml.findall("main:si", namespace):
                text_parts = [node.text or "" for node in item.findall(".//main:t", namespace)]
                shared_strings.append("".join(text_parts))

    return rel_map, sheet_path, sheet_xml, build_shared_string_lookup(shared_strings)


def build_shared_string_lookup(shared_strings: list[str]) -> dict[str, str]:
    return {str(index): value for index, value in enumerate(shared_strings)}


def get_cell_value(cell: ET.Element, shared_strings: dict[str, str]) -> str:
    cell_type = cell.attrib.get("t")
    value_node = cell.find(f"{{{XML_NS}}}v")
    if cell_type == "s" and value_node is not None:
        return shared_strings.get(value_node.text or "", "")
    if cell_type == "inlineStr":
        text_node = cell.find(f".//{{{XML_NS}}}t")
        return text_node.text if text_node is not None else ""
    return value_node.text if value_node is not None else ""


def load_template_rows_by_kind(template_path: Path) -> dict[str, dict[str, str]]:
    _, _, sheet_xml, shared_strings = load_template_workbook_details(template_path)
    namespace = {"main": XML_NS}

    template_rows: dict[str, dict[str, str]] = {}
    for row in sheet_xml.findall("main:sheetData/main:row", namespace):
        row_number = int(row.attrib.get("r", "0"))
        if row_number < 3:
            continue

        row_values: dict[str, str] = {}
        for cell in row.findall("main:c", namespace):
            reference = cell.attrib.get("r", "")
            column_letters = "".join(character for character in reference if character.isalpha())
            row_values[column_letters] = get_cell_value(cell, shared_strings).strip()

        kind = row_values.get("A", "").upper()
        if kind:
            template_rows[kind] = row_values

    return template_rows


def get_wrong_weight_flag(
    kind_of_jeans: str,
    order_weight_kg: str,
    template_rows_by_kind: dict[str, dict[str, str]],
) -> str:
    template_key = TEMPLATE_KIND_MAP.get(kind_of_jeans)
    if not template_key:
        return ""

    try:
        order_weight_grams = float(order_weight_kg) * 1000
    except ValueError:
        return ""

    template_row = template_rows_by_kind.get(template_key)
    if not template_row:
        return ""

    try:
        template_weight_grams = float(template_row.get("E", ""))
    except ValueError:
        return ""

    return "WRONG WEIGHT" if order_weight_grams > template_weight_grams else ""


def create_wrong_weight_file(
    template_path: Path,
    output_path: Path,
    context_rows: list[list[str]],
) -> int:
    _, sheet_path, sheet_xml, shared_strings = load_template_workbook_details(template_path)
    namespace = {"main": XML_NS}
    template_rows_by_kind = load_template_rows_by_kind(template_path)

    sheet_data = sheet_xml.find("main:sheetData", namespace)
    if sheet_data is None:
        raise ValueError("Template sheet data is missing.")

    header_rows = []
    for row in list(sheet_data.findall("main:row", namespace)):
        row_number = int(row.attrib.get("r", "0"))
        if row_number <= 2:
            header_rows.append(deepcopy(row))

    for row in list(sheet_data):
        sheet_data.remove(row)

    for row in header_rows:
        sheet_data.append(row)

    flagged_rows = [row for row in context_rows if len(row) > 9 and row[9] == "WRONG WEIGHT"]
    for output_index, row in enumerate(flagged_rows, start=3):
        kind = row[8]
        template_key = TEMPLATE_KIND_MAP.get(kind)
        template_values = template_rows_by_kind.get(template_key or "", {})
        order_weight_grams = convert_kg_to_grams(row[6])

        output_values = {
            "A": row[0],
            "B": template_values.get("B", ""),
            "C": template_values.get("C", ""),
            "D": template_values.get("D", ""),
            "E": template_values.get("E", ""),
            "F": row[3],
            "G": row[4],
            "H": row[5],
            "I": order_weight_grams,
            "J": template_values.get("J", ""),
            "K": template_values.get("K", ""),
            "L": template_values.get("L", ""),
            "M": template_values.get("M", ""),
        }
        sheet_data.append(build_sheet_row(output_index, output_values))

    dimension = sheet_xml.find("main:dimension", namespace)
    if dimension is not None:
        last_row = max(2, len(flagged_rows) + 2)
        dimension.set("ref", f"A1:M{last_row}")

    save_workbook_with_updated_sheet(template_path, output_path, sheet_path, sheet_xml)
    return len(flagged_rows)


def save_workbook_with_updated_sheet(
    template_path: Path,
    output_path: Path,
    target_sheet_path: str,
    updated_sheet_xml: ET.Element,
) -> None:
    updated_sheet_bytes = ET.tostring(updated_sheet_xml, encoding="utf-8", xml_declaration=True)

    with ZipFile(template_path, "r") as source_zip, ZipFile(output_path, "w", compression=ZIP_DEFLATED) as target_zip:
        for zip_info in source_zip.infolist():
            content = updated_sheet_bytes if zip_info.filename == target_sheet_path else source_zip.read(zip_info.filename)
            target_zip.writestr(zip_info, content)


def build_sheet_row(row_number: int, values_by_column: dict[str, str]) -> ET.Element:
    row_element = ET.Element(f"{{{XML_NS}}}row", {"r": str(row_number)})
    for column_letter, value in values_by_column.items():
        if value == "":
            continue

        cell_reference = f"{column_letter}{row_number}"
        if is_numeric_string(value):
            cell_element = ET.Element(f"{{{XML_NS}}}c", {"r": cell_reference})
            value_element = ET.SubElement(cell_element, f"{{{XML_NS}}}v")
            value_element.text = str(value)
        else:
            cell_element = ET.Element(f"{{{XML_NS}}}c", {"r": cell_reference, "t": "inlineStr"})
            is_element = ET.SubElement(cell_element, f"{{{XML_NS}}}is")
            text_element = ET.SubElement(is_element, f"{{{XML_NS}}}t")
            text_element.text = str(value)
        row_element.append(cell_element)

    return row_element


def is_numeric_string(value: str) -> bool:
    try:
        float(value)
        return True
    except ValueError:
        return False


def convert_kg_to_grams(weight_kg: str) -> str:
    try:
        grams = float(weight_kg) * 1000
    except ValueError:
        return ""

    return str(int(grams)) if grams.is_integer() else str(grams)


def excel_column_name(column_number: int) -> str:
    name = []
    while column_number > 0:
        column_number, remainder = divmod(column_number - 1, 26)
        name.append(chr(65 + remainder))
    return "".join(reversed(name))


def write_excel(output_path: Path, rows: list[list[str]]) -> None:
    all_rows = [
        [column_name for column_name, _ in SOURCE_COLUMNS] + ["KIND_OF_JEANS", "WRONG_WEIGHT_FLAG"],
        *rows,
    ]

    worksheet_rows = []
    for row_index, row in enumerate(all_rows, start=1):
        cells = []
        for column_index, value in enumerate(row, start=1):
            cell_ref = f"{excel_column_name(column_index)}{row_index}"
            cell_value = escape(str(value))
            cells.append(
                f'<c r="{cell_ref}" t="inlineStr"><is><t>{cell_value}</t></is></c>'
            )
        worksheet_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    worksheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData>"
        f'{"".join(worksheet_rows)}'
        "</sheetData>"
        "</worksheet>"
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets>"
        '<sheet name="Current Context" sheetId="1" r:id="rId1"/>'
        "</sheets>"
        "</workbook>"
    )

    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        "</Relationships>"
    )

    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>"
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )

    with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as workbook_zip:
        workbook_zip.writestr("[Content_Types].xml", content_types_xml)
        workbook_zip.writestr("_rels/.rels", root_rels_xml)
        workbook_zip.writestr("xl/workbook.xml", workbook_xml)
        workbook_zip.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        workbook_zip.writestr("xl/worksheets/sheet1.xml", worksheet_xml)
        workbook_zip.writestr("xl/styles.xml", styles_xml)


class WeightCorrectorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Weight Corrector")
        self.root.geometry("640x260")
        self.root.resizable(False, False)

        self.selected_file = tk.StringVar()
        self.status_text = tk.StringVar(value="Choose an Order CSV file to begin.")

        self.build_ui()

    def build_ui(self) -> None:
        container = tk.Frame(self.root, padx=20, pady=20)
        container.pack(fill="both", expand=True)

        title = tk.Label(
            container,
            text="Weight Corrector",
            font=("Segoe UI", 18, "bold"),
        )
        title.pack(anchor="w", pady=(0, 8))

        subtitle = tk.Label(
            container,
            text="Upload an Order CSV file and the app will automatically create current_context.xlsx and WRONG_WEIGHT_FILE.xlsx next to this script.",
            font=("Segoe UI", 10),
            wraplength=580,
            justify="left",
        )
        subtitle.pack(anchor="w", pady=(0, 16))

        file_row = tk.Frame(container)
        file_row.pack(fill="x", pady=(0, 12))

        file_entry = tk.Entry(
            file_row,
            textvariable=self.selected_file,
            font=("Segoe UI", 10),
            state="readonly",
            readonlybackground="white",
        )
        file_entry.pack(side="left", fill="x", expand=True)

        browse_button = tk.Button(
            file_row,
            text="Upload Order CSV",
            font=("Segoe UI", 10, "bold"),
            command=self.process_order_csv,
            padx=14,
            pady=8,
        )
        browse_button.pack(side="left", padx=(10, 0))

        status_label = tk.Label(
            container,
            textvariable=self.status_text,
            font=("Segoe UI", 10),
            fg="#1f3a5f",
            wraplength=580,
            justify="left",
        )
        status_label.pack(anchor="w")

    def process_order_csv(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Order CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not file_path:
            self.status_text.set("No file selected.")
            return

        self.selected_file.set(file_path)
        input_path = Path(file_path)
        if not input_path.exists():
            messagebox.showerror("File not found", f"Could not find:\n{input_path}")
            return

        template_path = Path(__file__).with_name(DEFAULT_TEMPLATE_NAME)
        if not template_path.exists():
            messagebox.showerror(
                "Template not found",
                f"Could not find the fixed template file:\n{template_path}",
            )
            return

        script_directory = Path(__file__).resolve().parent
        context_output_path = script_directory / CURRENT_CONTEXT_NAME
        wrong_weight_output_path = script_directory / WRONG_WEIGHT_FILE_NAME

        try:
            rows = read_selected_columns(input_path)
            write_excel(context_output_path, rows)
            flagged_count = create_wrong_weight_file(template_path, wrong_weight_output_path, rows)
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to create the output files.\n\n{exc}")
            self.status_text.set("Something went wrong while creating the output files.")
            return

        self.status_text.set(f"Created: {context_output_path.name} and {wrong_weight_output_path.name}")
        messagebox.showinfo(
            "Success",
            "Files created successfully.\n\n"
            f"Rows exported to {CURRENT_CONTEXT_NAME}: {len(rows)}\n"
            f"Wrong weight rows in {WRONG_WEIGHT_FILE_NAME}: {flagged_count}\n\n"
            f"Saved in:\n{script_directory}",
        )


def main() -> None:
    root = tk.Tk()
    app = WeightCorrectorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
