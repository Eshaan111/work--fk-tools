import html
import json
import re
import sys
import zipfile
from pathlib import Path
import pandas as pd
import numpy as np
from statistics import mode, StatisticsError
from xml.etree import ElementTree as ET

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QTableWidget, QTableWidgetItem, QPushButton,
    QFileDialog, QLineEdit, QSlider, QMessageBox, QTextEdit
)
from PyQt5.QtCore import Qt

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


class Dashboard(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PRO Settlement Dashboard")
        self.setGeometry(100, 100, 1300, 750)

        self.df = pd.DataFrame()
        self.selected_range = None
        self.offer_mode = False
        self.col_sku = ""
        self.col_title = ""
        self.col_settlement = ""
        self.file_type = "xlsx"
        self.available_modes = {"normal": False, "offer": False}
        self.size_override_path = Path(__file__).with_name("size_overrides.json")
        self.flag_config_path = Path(__file__).with_name("flag_rules.xlsx")
        self.size_overrides = {}
        self.size_values = ["26", "28", "30", "32", "34", "36"]
        self.default_title_flag_keywords = ["Dark Blue"]
        self.default_sku_flag_keywords = []
        self.default_inactive_sizes = ["26", "36"]
        self.flag_title_keywords = list(self.default_title_flag_keywords)
        self.flag_sku_keywords = list(self.default_sku_flag_keywords)
        self.flag_inactive_sizes = list(self.default_inactive_sizes)
        self.undo_state = None
        self.undo_label = ""
        self.change_log_entries = []
        self.drag_start = None
        self.hover_cid = None
        self.press_cid = None
        self.release_cid = None

        # -------- UI --------
        main_widget = QWidget()
        main_widget.setObjectName("mainPanel")
        self.setCentralWidget(main_widget)
        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(18, 18, 18, 18)
        self.layout.setSpacing(12)
        main_widget.setLayout(self.layout)
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        self.compute_btn = QPushButton("Compute Discount")
        self.compute_btn.clicked.connect(self.compute_discount_only)

        self.decision_btn = QPushButton("Apply Decision")
        self.decision_btn.clicked.connect(self.apply_decision_only)

        btn_layout.addWidget(self.compute_btn)
        btn_layout.addWidget(self.decision_btn)

        self.layout.addLayout(btn_layout)
        self.offer_controls = [self.compute_btn, self.decision_btn]

        # -------- TOP --------
        top = QHBoxLayout()
        top.setSpacing(10)

        self.load_btn = QPushButton("Load Excel")
        self.load_btn.clicked.connect(self.load_file)

        self.export_btn = QPushButton("Export File")
        self.export_btn.clicked.connect(self.export_file)

        self.undo_btn = QPushButton("Undo Last Move")
        self.undo_btn.clicked.connect(self.undo_last_move)
        self.undo_btn.setEnabled(False)

        self.mode_btn = QPushButton("Offer Mode: OFF")
        self.mode_btn.clicked.connect(self.toggle_mode)

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search SKU or Title...")
        self.search_box.textChanged.connect(self.update_dashboard)

        top.addWidget(self.load_btn)
        top.addWidget(self.export_btn)
        top.addWidget(self.undo_btn)
        top.addWidget(self.mode_btn)
        top.addWidget(self.search_box)

        self.layout.addLayout(top)

        # -------- OFFER INPUTS --------
        offer_layout = QHBoxLayout()
        offer_layout.setSpacing(10)

        self.y_input = QLineEdit()
        self.y_input.setPlaceholderText("y%")

        self.x_input = QLineEdit()
        self.x_input.setPlaceholderText("x%")

        self.cap_input = QLineEdit()
        self.cap_input.setPlaceholderText("Cap ₹")

        offer_layout.addWidget(self.y_input)
        offer_layout.addWidget(self.x_input)
        offer_layout.addWidget(self.cap_input)

        self.layout.addLayout(offer_layout)
        self.offer_controls.extend([self.y_input, self.x_input, self.cap_input])

        # -------- THRESHOLDS --------
        self.thresholds = {}
        threshold_layout = QHBoxLayout()
        threshold_layout.setSpacing(10)

        for j in ["ICE","BEIGE","WHITE","BLACK-BAGGY","BLACK-PLAIN","MIX"]:
            inp = QLineEdit()
            inp.setPlaceholderText(j)
            self.thresholds[j] = inp
            threshold_layout.addWidget(inp)

        self.layout.addLayout(threshold_layout)
        self.offer_controls.extend(self.thresholds.values())

        # -------- FILTER --------
        filters = QHBoxLayout()
        filters.setSpacing(10)

        self.listing_filter = QComboBox()
        self.jeans_filter = QComboBox()
        self.size_filter = QComboBox()
        self.status_filter = QComboBox()
        self.slider = QSlider(Qt.Horizontal)
        self.listing_filter.currentIndexChanged.connect(self.update_dashboard)
        self.jeans_filter.currentIndexChanged.connect(self.update_dashboard)
        self.size_filter.currentIndexChanged.connect(self.update_dashboard)
        self.status_filter.currentIndexChanged.connect(self.update_dashboard)
        self.slider.valueChanged.connect(self.update_dashboard)

        filters.addWidget(QLabel("Listing"))
        filters.addWidget(self.listing_filter)
        filters.addWidget(QLabel("Jeans"))
        filters.addWidget(self.jeans_filter)
        filters.addWidget(QLabel("Size"))
        filters.addWidget(self.size_filter)
        filters.addWidget(QLabel("Status"))
        filters.addWidget(self.status_filter)
        filters.addWidget(QLabel("Settlement Max"))
        filters.addWidget(self.slider)

        self.layout.addLayout(filters)

        # -------- DASHBOARD META --------
        meta_row = QHBoxLayout()
        meta_row.setSpacing(10)
        self.row_count_label = QLabel("Loaded: 0 | Visible: 0 | Export: 0")
        self.row_count_label.setObjectName("metaLabel")
        self.account_label = QLabel("Account: Unknown")
        self.account_label.setObjectName("metaLabel")
        meta_row.addWidget(self.row_count_label, 1)
        meta_row.addWidget(self.account_label, 0)
        self.layout.addLayout(meta_row)

        status_matrix_row = QHBoxLayout()
        status_matrix_row.setSpacing(10)
        self.status_matrix_label = QLabel("Size Status Matrix\nACTIVE: -\nINACTIVE: -")
        self.status_matrix_label.setObjectName("metaLabel")
        self.status_matrix_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.status_matrix_label.setMinimumHeight(72)
        status_matrix_row.addWidget(self.status_matrix_label)
        self.layout.addLayout(status_matrix_row)

        # -------- SIZE OVERRIDES --------
        size_row = QHBoxLayout()
        size_row.setSpacing(10)

        self.size_sku_input = QLineEdit()
        self.size_sku_input.setPlaceholderText("SKU for size override")
        self.size_override_combo = QComboBox()
        self.size_override_combo.addItems(self.size_values)
        self.save_size_btn = QPushButton("Save Size")
        self.save_size_btn.clicked.connect(self.save_size_override)
        self.download_undetected_btn = QPushButton("Download Undetected")
        self.download_undetected_btn.clicked.connect(self.download_undetected_sizes)
        self.upload_sizes_btn = QPushButton("Upload Sizes")
        self.upload_sizes_btn.clicked.connect(self.upload_size_sheet)

        size_row.addWidget(self.size_sku_input)
        size_row.addWidget(self.size_override_combo)
        size_row.addWidget(self.save_size_btn)
        size_row.addWidget(self.download_undetected_btn)
        size_row.addWidget(self.upload_sizes_btn)

        self.layout.addLayout(size_row)

        # -------- FREEZE --------
        freeze = QHBoxLayout()
        freeze.setSpacing(10)

        self.freeze_col = QComboBox()
        self.freeze_val = QLineEdit()
        freeze_btn = QPushButton("Freeze")
        freeze_btn.clicked.connect(self.apply_freeze)
        unfreeze_btn = QPushButton("Unfreeze All")
        unfreeze_btn.clicked.connect(self.unfreeze_all)

        freeze.addWidget(self.freeze_col)
        freeze.addWidget(self.freeze_val)
        freeze.addWidget(freeze_btn)
        freeze.addWidget(unfreeze_btn)

        self.layout.addLayout(freeze)

        # -------- STATS + BULK EDIT --------
        stats = QHBoxLayout()
        stats.setSpacing(10)

        self.stats_label = QLabel("No Selection")
        self.edit_mode = QComboBox()
        self.edit_mode.addItems(["Add", "Multiply", "Replace"])
        self.edit_value = QLineEdit()
        self.edit_value.setPlaceholderText("Value")
        self.edit_cap_mode = QComboBox()
        self.edit_cap_mode.addItems(["Min", "Max"])
        self.edit_cap_value = QLineEdit()
        self.edit_cap_value.setPlaceholderText("Cap value")
        apply_btn = QPushButton("Apply")
        apply_btn.clicked.connect(self.apply_bulk_edit)
        self.status_change_combo = QComboBox()
        self.status_change_combo.addItems(["ACTIVE", "INACTIVE"])
        self.status_apply_btn = QPushButton("Set Status")
        self.status_apply_btn.clicked.connect(self.apply_status_change)

        stats.addWidget(self.stats_label)
        stats.addWidget(self.edit_mode)
        stats.addWidget(self.edit_value)
        stats.addWidget(self.edit_cap_mode)
        stats.addWidget(self.edit_cap_value)
        stats.addWidget(apply_btn)
        stats.addWidget(self.status_change_combo)
        stats.addWidget(self.status_apply_btn)

        self.layout.addLayout(stats)

        # -------- CHANGE LOG --------
        self.change_log = QTextEdit()
        self.change_log.setObjectName("changeLog")
        self.change_log.setReadOnly(True)
        self.change_log.setPlaceholderText("Changes made will appear here...")
        self.change_log.setMaximumHeight(180)
        self.layout.addWidget(self.change_log)

        # -------- RESET --------
        reset_btn = QPushButton("Reset Selection")
        reset_btn.clicked.connect(self.reset_selection)
        self.layout.addWidget(reset_btn)

        # -------- TABLE --------
        self.table = QTableWidget()
        self.table.itemChanged.connect(self.handle_edit)
        self.layout.addWidget(self.table)

        # -------- GRAPH --------
        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)
        self.layout.addWidget(self.canvas)
        self.apply_dashboard_style()
        self.refresh_mode_ui()
        self.load_flag_config()

    # ---------------- MODE TOGGLE ----------------
    def toggle_mode(self):
        if self.df.empty:
            self.offer_mode = not self.offer_mode
            self.refresh_mode_ui()
            return

        target_mode = not self.offer_mode
        if target_mode and not self.available_modes["offer"]:
            QMessageBox.warning(self, "Error", "This file does not support Offer Mode")
            return

        if not target_mode and not self.available_modes["normal"]:
            QMessageBox.warning(self, "Error", "This file does not support Normal Mode")
            return

        self.offer_mode = target_mode
        self.apply_mode_config()
        self.update_dashboard()

    def refresh_mode_ui(self):
        self.mode_btn.setText("Offer Mode: ON" if self.offer_mode else "Offer Mode: OFF")
        for widget in self.offer_controls:
            widget.setVisible(self.offer_mode)

    def apply_dashboard_style(self):
        self.setStyleSheet("""
            QWidget#mainPanel {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #f7f1e8, stop:0.45 #f4efe8, stop:1 #ede3d3);
            }
            QLabel {
                color: #33261c;
                font-size: 12px;
            }
            QLabel#metaLabel {
                background: rgba(255, 252, 247, 0.92);
                border: 1px solid #dccbb7;
                border-radius: 10px;
                padding: 8px 12px;
                font-size: 12px;
                font-weight: 600;
                color: #5c3d2e;
            }
            QLineEdit, QComboBox, QTextEdit {
                background: #fffdf9;
                border: 1px solid #d8c8b5;
                border-radius: 10px;
                padding: 7px 10px;
                color: #2f241b;
                selection-background-color: #c96f3b;
            }
            QLineEdit:focus, QComboBox:focus, QTextEdit:focus {
                border: 1px solid #c96f3b;
            }
            QPushButton {
                background: #5f7c6c;
                color: white;
                border: none;
                border-radius: 10px;
                padding: 8px 14px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #4f6a5b;
            }
            QPushButton:pressed {
                background: #405649;
            }
            QPushButton:disabled {
                background: #b7b0a8;
                color: #f6f2ec;
            }
            QTableWidget {
                background: rgba(255, 253, 249, 0.97);
                alternate-background-color: #f5ede2;
                border: 1px solid #d8c8b5;
                border-radius: 12px;
                gridline-color: #e5d8c8;
                color: #2f241b;
            }
            QHeaderView::section {
                background: #e7d6c4;
                color: #4f3425;
                border: none;
                border-right: 1px solid #d4c1ab;
                padding: 8px;
                font-weight: 700;
            }
            QTextEdit#changeLog {
                background: #fffaf3;
                border: 1px solid #d8c8b5;
                border-radius: 12px;
                padding: 6px;
            }
            QSlider::groove:horizontal {
                border: 0;
                height: 8px;
                background: #dbcab7;
                border-radius: 4px;
            }
            QSlider::handle:horizontal {
                background: #c96f3b;
                width: 18px;
                margin: -6px 0;
                border-radius: 9px;
            }
        """)
        self.table.setAlternatingRowColors(True)

    def apply_mode_config(self):
        if self.offer_mode:
            self.col_sku = "SKU ID"
            self.col_title = "FSN"
            self.col_settlement = "Selling Price(Rs)"
        else:
            self.col_sku = "Seller SKU Id"
            self.col_title = "Product Title"
            self.col_settlement = "Bank Settlement"

        self.refresh_mode_ui()

        if not self.df.empty:
            self.clean_numeric(self.col_settlement)
            self.apply_classification()
            self.setup_filters()

    def load_size_overrides(self):
        if not self.size_override_path.exists():
            self.size_overrides = {}
            return

        try:
            self.size_overrides = json.loads(self.size_override_path.read_text(encoding="utf-8"))
        except Exception:
            self.size_overrides = {}

    def persist_size_overrides(self):
        self.size_override_path.write_text(
            json.dumps(self.size_overrides, indent=2, sort_keys=True),
            encoding="utf-8"
        )

    def excel_column_name(self, index):
        result = ""
        index = int(index)
        while index > 0:
            index, remainder = divmod(index - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def excel_column_index(self, col_name):
        result = 0
        for char in str(col_name).upper():
            if "A" <= char <= "Z":
                result = result * 26 + (ord(char) - 64)
        return result

    def write_simple_xlsx(self, path, rows):
        sheet_rows = []
        for row_idx, row_values in enumerate(rows, start=1):
            cells = []
            for col_idx, value in enumerate(row_values, start=1):
                if value is None or str(value) == "":
                    continue
                cell_ref = f"{self.excel_column_name(col_idx)}{row_idx}"
                cell_text = html.escape(str(value))
                cells.append(
                    f'<c r="{cell_ref}" t="inlineStr"><is><t>{cell_text}</t></is></c>'
                )
            sheet_rows.append(f'<row r="{row_idx}">{"".join(cells)}</row>')

        sheet_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<sheetData>{"".join(sheet_rows)}</sheetData>'
            '</worksheet>'
        )
        workbook_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Flags" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>'
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
            '</Types>'
        )
        root_rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/>'
            '</Relationships>'
        )
        workbook_rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/>'
            '</Relationships>'
        )

        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", content_types_xml)
            zf.writestr("_rels/.rels", root_rels_xml)
            zf.writestr("xl/workbook.xml", workbook_xml)
            zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
            zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)

    def read_simple_xlsx(self, path):
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        rows = []

        with zipfile.ZipFile(path, "r") as zf:
            shared_strings = []
            if "xl/sharedStrings.xml" in zf.namelist():
                shared_xml = ET.fromstring(zf.read("xl/sharedStrings.xml"))
                for si in shared_xml.findall(".//main:si", ns):
                    text_nodes = si.findall(".//main:t", ns)
                    shared_strings.append("".join(node.text or "" for node in text_nodes))
            sheet_xml = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))

        for row in sheet_xml.findall(".//main:sheetData/main:row", ns):
            values = []
            current_col = 1
            for cell in row.findall("main:c", ns):
                ref = cell.attrib.get("r", "")
                col_name = "".join(ch for ch in ref if ch.isalpha())
                col_index = self.excel_column_index(col_name) or current_col

                while current_col < col_index:
                    values.append("")
                    current_col += 1

                cell_type = cell.attrib.get("t", "")
                cell_value = ""
                if cell_type == "inlineStr":
                    text_nodes = cell.findall(".//main:t", ns)
                    cell_value = "".join(node.text or "" for node in text_nodes)
                elif cell_type == "s":
                    value_node = cell.find("main:v", ns)
                    if value_node is not None and value_node.text is not None:
                        try:
                            shared_index = int(value_node.text)
                            if 0 <= shared_index < len(shared_strings):
                                cell_value = shared_strings[shared_index]
                        except Exception:
                            cell_value = ""
                else:
                    value_node = cell.find("main:v", ns)
                    cell_value = value_node.text if value_node is not None else ""

                values.append(cell_value)
                current_col = col_index + 1

            rows.append(values)

        return rows

    def create_default_flag_config(self):
        rows = [
            ["Title Keyword", *self.default_title_flag_keywords],
            ["SKU", *self.default_sku_flag_keywords],
            ["Size", *self.default_inactive_sizes],
        ]
        self.write_simple_xlsx(self.flag_config_path, rows)

    def load_flag_config(self):
        if not self.flag_config_path.exists():
            self.create_default_flag_config()

        self.flag_title_keywords = list(self.default_title_flag_keywords)
        self.flag_sku_keywords = list(self.default_sku_flag_keywords)
        self.flag_inactive_sizes = list(self.default_inactive_sizes)

        try:
            config_rows = self.read_simple_xlsx(self.flag_config_path)
        except Exception:
            return

        title_keywords = []
        sku_keywords = []
        inactive_sizes = []

        for row in config_rows:
            if not row:
                continue

            row_key = str(row[0]).strip().lower()
            values = [
                str(value).strip()
                for value in row[1:]
                if str(value).strip()
            ]

            if row_key == "title keyword" and values:
                title_keywords = values
            elif row_key == "sku" and values:
                sku_keywords = values
            elif row_key == "size" and values:
                inactive_sizes = values

        if title_keywords:
            self.flag_title_keywords = title_keywords
        if sku_keywords:
            self.flag_sku_keywords = sku_keywords
        if inactive_sizes:
            self.flag_inactive_sizes = inactive_sizes

    def detect_title_flag_matches(self, title_text):
        title = str(title_text).strip()
        if not title:
            return []

        title_lower = title.lower()
        return [
            keyword for keyword in self.flag_title_keywords
            if str(keyword).strip() and str(keyword).strip().lower() in title_lower
        ]

    def detect_sku_flag_matches(self, sku_text):
        sku = str(sku_text).strip()
        if not sku:
            return []

        sku_lower = sku.lower()
        return [
            keyword for keyword in self.flag_sku_keywords
            if str(keyword).strip() and str(keyword).strip().lower() in sku_lower
        ]

    def ensure_flag_exemption_column(self):
        if "__flag_exemptions" not in self.df.columns:
            self.df["__flag_exemptions"] = ""

    def parse_flag_exemptions(self, raw_value):
        text = str(raw_value).strip()
        if not text:
            return set()
        try:
            parsed = json.loads(text)
            if isinstance(parsed, list):
                return {str(item).strip() for item in parsed if str(item).strip()}
        except Exception:
            pass
        return {part.strip() for part in text.split("|") if part.strip()}

    def serialize_flag_exemptions(self, values):
        clean_values = sorted({str(value).strip() for value in values if str(value).strip()})
        return json.dumps(clean_values)

    def build_flag_details(self, mask=None):
        if self.df.empty:
            return pd.DataFrame(index=self.df.index)

        if mask is None:
            mask = pd.Series(True, index=self.df.index, dtype=bool)
        else:
            mask = mask.reindex(self.df.index, fill_value=False)

        self.ensure_flag_exemption_column()

        title_series = self.df[self.col_title].astype(str) if self.col_title in self.df.columns else pd.Series("", index=self.df.index)
        sku_series = self.df[self.col_sku].astype(str) if self.col_sku in self.df.columns else pd.Series("", index=self.df.index)
        size_series = self.df["Size"].astype(str).str.strip()
        title_matches = title_series.apply(self.detect_title_flag_matches)
        sku_matches = sku_series.apply(self.detect_sku_flag_matches)
        exemption_sets = self.df["__flag_exemptions"].apply(self.parse_flag_exemptions)

        size_reason = pd.Series("", index=self.df.index, dtype=object)
        title_reason = pd.Series("", index=self.df.index, dtype=object)
        sku_reason = pd.Series("", index=self.df.index, dtype=object)
        auto_flag = pd.Series("", index=self.df.index, dtype=object)

        for idx in self.df.index[mask]:
            reasons = []
            size_value = size_series.loc[idx]
            title_keywords = title_matches.loc[idx]
            sku_keywords = sku_matches.loc[idx]
            exemptions = exemption_sets.loc[idx]

            if size_value in self.flag_inactive_sizes and f"size:{size_value}" not in exemptions:
                size_reason.loc[idx] = size_value
                reasons.append(f"SIZE: {size_value}")

            active_title_keywords = [
                keyword for keyword in title_keywords
                if f"title:{keyword.lower()}" not in exemptions
            ]
            if active_title_keywords:
                title_reason.loc[idx] = " | ".join(active_title_keywords)
                reasons.append("TITLE: " + " | ".join(active_title_keywords))

            active_sku_keywords = [
                keyword for keyword in sku_keywords
                if f"sku:{keyword.lower()}" not in exemptions
            ]
            if active_sku_keywords:
                sku_reason.loc[idx] = " | ".join(active_sku_keywords)
                reasons.append("SKU: " + " | ".join(active_sku_keywords))

            auto_flag.loc[idx] = " | ".join(reasons)

        return pd.DataFrame({
            "__flag_size_reason": size_reason,
            "__flag_title_reason": title_reason,
            "__flag_sku_reason": sku_reason,
            "__auto_flag": auto_flag,
        }, index=self.df.index)

    def build_flag_summary_lines(self, flagged_df):
        if flagged_df is None or flagged_df.empty or "Jeans Type" not in flagged_df.columns:
            return []

        flagged_rows = flagged_df[["Jeans Type", "__flag_size_reason", "__flag_title_reason", "__flag_sku_reason"]].copy()
        flagged_rows["Jeans Type"] = flagged_rows["Jeans Type"].astype(str).str.strip()
        flagged_rows["__flag_size_reason"] = flagged_rows["__flag_size_reason"].astype(str).str.strip()
        flagged_rows["__flag_title_reason"] = flagged_rows["__flag_title_reason"].astype(str).str.strip()
        flagged_rows["__flag_sku_reason"] = flagged_rows["__flag_sku_reason"].astype(str).str.strip()

        unique_lines = set()
        for _, row in flagged_rows.iterrows():
            kind = row["Jeans Type"] or "UNKNOWN"
            if row["__flag_size_reason"]:
                unique_lines.add(f"{kind} : {row['__flag_size_reason']}")
            if row["__flag_title_reason"]:
                for keyword in row["__flag_title_reason"].split(" | "):
                    keyword = keyword.strip()
                    if keyword:
                        unique_lines.add(f"{kind} : {keyword}")
            if row["__flag_sku_reason"]:
                for keyword in row["__flag_sku_reason"].split(" | "):
                    keyword = keyword.strip()
                    if keyword:
                        unique_lines.add(f"{kind} : {keyword}")

        return sorted(unique_lines)

    def resolve_activation_mask(self, mask, action_label):
        if self.df.empty or "Listing Status" not in self.df.columns:
            return mask

        details = self.build_flag_details(mask)
        flagged_mask = mask & (details["__auto_flag"].astype(str).str.strip() != "")
        if not flagged_mask.any():
            return mask

        flagged_df = self.df.loc[flagged_mask].copy()
        flagged_df["__flag_size_reason"] = details.loc[flagged_mask, "__flag_size_reason"]
        flagged_df["__flag_title_reason"] = details.loc[flagged_mask, "__flag_title_reason"]
        flagged_df["__flag_sku_reason"] = details.loc[flagged_mask, "__flag_sku_reason"]
        lines = self.build_flag_summary_lines(flagged_df)
        if not lines:
            return mask

        popup = QMessageBox(self)
        popup.setIcon(QMessageBox.Warning)
        popup.setWindowTitle("Flag Override")
        popup.setText(
            "This action will make flagged rows ACTIVE:\n\n"
            + "\n".join(lines)
            + "\n\nChoose what to do with these flagged rows."
        )
        activate_button = popup.addButton("Activate Flagged Rows", QMessageBox.AcceptRole)
        keep_inactive_button = popup.addButton("Keep Flagged Rows Inactive", QMessageBox.RejectRole)
        popup.setDefaultButton(keep_inactive_button)
        popup.exec_()

        if popup.clickedButton() == activate_button:
            self.ensure_flag_exemption_column()
            for idx in self.df.index[flagged_mask]:
                exemptions = self.parse_flag_exemptions(self.df.at[idx, "__flag_exemptions"])
                size_reason = str(details.at[idx, "__flag_size_reason"]).strip()
                title_reason = str(details.at[idx, "__flag_title_reason"]).strip()
                sku_reason = str(details.at[idx, "__flag_sku_reason"]).strip()
                if size_reason:
                    exemptions.add(f"size:{size_reason}")
                if title_reason:
                    for keyword in title_reason.split(" | "):
                        keyword = keyword.strip()
                        if keyword:
                            exemptions.add(f"title:{keyword.lower()}")
                if sku_reason:
                    for keyword in sku_reason.split(" | "):
                        keyword = keyword.strip()
                        if keyword:
                            exemptions.add(f"sku:{keyword.lower()}")
                self.df.at[idx, "__flag_exemptions"] = self.serialize_flag_exemptions(exemptions)
            return mask

        return mask & ~flagged_mask

    def save_undo_state(self, label):
        if self.df.empty:
            return

        self.undo_state = {
            "df": self.df.copy(deep=True),
            "size_overrides": dict(self.size_overrides),
        }
        self.undo_label = label
        self.undo_btn.setEnabled(True)
        self.undo_btn.setText(f"Undo: {label}")

    def undo_last_move(self):
        if not self.undo_state:
            return

        self.df = self.undo_state["df"].copy(deep=True)
        self.size_overrides = dict(self.undo_state["size_overrides"])
        self.persist_size_overrides()
        self.undo_state = None
        self.undo_label = ""
        self.undo_btn.setEnabled(False)
        self.undo_btn.setText("Undo Last Move")

        if not self.df.empty:
            self.apply_classification()
            self.setup_filters()
        self.update_dashboard()

    def format_min_max(self, values):
        if values is None or len(values) == 0:
            return "NA / NA"
        return f"{float(values.min()):.2f} / {float(values.max()):.2f}"

    def summarize_context_df(self, df_subset):
        if df_subset is None or df_subset.empty:
            return ""

        def uniq(col):
            if col not in df_subset.columns:
                return ""
            values = df_subset[col].dropna().astype(str).unique().tolist()
            return "/".join(values[:4]) if values else ""

        parts = []
        jeans = uniq("Jeans Type")
        listing = uniq("Listing Type")
        size = uniq("Size")
        status = uniq("Listing Status")

        if jeans:
            parts.append(f"Jeans: {jeans}")
        if listing:
            parts.append(f"Listing: {listing}")
        if size:
            parts.append(f"Size: {size}")
        if status:
            parts.append(f"Status: {status}")

        return " | ".join(parts)

    def render_change_log(self):
        if not self.change_log_entries:
            self.change_log.clear()
            return

        cards = []
        for entry in self.change_log_entries[-12:]:
            cards.append(f"""
                <div style="
                    background:#fffdf9;
                    border:1px solid #e0cfbc;
                    border-left:6px solid #c96f3b;
                    border-radius:12px;
                    padding:10px 12px;
                    margin:0 0 8px 0;">
                    <div style="font-weight:700; color:#5a3625; font-size:13px;">
                        {html.escape(entry["action"])}
                    </div>
                    <div style="color:#6e4c39; font-size:11px; margin-top:4px;">
                        Rows changed: {entry["rows"]}
                    </div>
                    <div style="color:#3d2c22; font-size:11px; margin-top:6px;">
                        <b>Before Min/Max:</b> {html.escape(entry["before"])}
                    </div>
                    <div style="color:#3d2c22; font-size:11px; margin-top:2px;">
                        <b>After Min/Max:</b> {html.escape(entry["after"])}
                    </div>
                    {f'<div style="color:#7a5a46; font-size:11px; margin-top:6px;">{html.escape(entry["extra"])}</div>' if entry["extra"] else ''}
                    {f'<div style="color:#7a5a46; font-size:11px; margin-top:4px;">{html.escape(entry["context"])}</div>' if entry["context"] else ''}
                </div>
            """)

        self.change_log.setHtml(
            "<div style='font-family:Segoe UI, sans-serif;'>" + "".join(cards) + "</div>"
        )
        self.change_log.verticalScrollBar().setValue(self.change_log.verticalScrollBar().maximum())

    def log_change(self, action, before_values, after_values, affected_rows, extra="", context_df=None):
        before_text = self.format_min_max(before_values)
        after_text = self.format_min_max(after_values)
        context = self.summarize_context_df(context_df)
        self.change_log_entries.append({
            "action": action,
            "rows": affected_rows,
            "before": before_text,
            "after": after_text,
            "extra": extra,
            "context": context,
        })
        self.render_change_log()

    def detect_size_from_sku(self, sku):
        sku_text = str(sku).strip()
        if not sku_text:
            return "UNDETECTED"

        if sku_text in self.size_overrides:
            return self.size_overrides[sku_text]

        if re.search(r"_39$", sku_text, flags=re.IGNORECASE):
            return "32"

        for size in self.size_values:
            patterns = [
                rf"-{size}-",
                rf"_{size}_",
                rf"_{size}$",
            ]
            if any(re.search(pattern, sku_text, flags=re.IGNORECASE) for pattern in patterns):
                return size

        return "UNDETECTED"

    def apply_size_flags(self, mask=None):
        if self.df.empty:
            return self.df.iloc[0:0].copy()

        if mask is None:
            mask = pd.Series(True, index=self.df.index, dtype=bool)
        else:
            mask = mask.reindex(self.df.index, fill_value=False)

        if "Auto Flag" not in self.df.columns:
            self.df["Auto Flag"] = ""

        details = self.build_flag_details(mask)
        self.df.loc[mask, "Auto Flag"] = details.loc[mask, "__auto_flag"]

        flagged_mask = mask & (details["__auto_flag"].astype(str).str.strip() != "")
        clear_flag_mask = mask & ~flagged_mask
        self.df.loc[clear_flag_mask, "Auto Flag"] = ""

        if "Listing Status" in self.df.columns:
            self.df.loc[flagged_mask, "Listing Status"] = "INACTIVE"

        flagged_df = self.df.loc[flagged_mask].copy()
        flagged_df["__flag_size_reason"] = details.loc[flagged_mask, "__flag_size_reason"]
        flagged_df["__flag_title_reason"] = details.loc[flagged_mask, "__flag_title_reason"]
        flagged_df["__flag_sku_reason"] = details.loc[flagged_mask, "__flag_sku_reason"]
        return flagged_df

    def show_flag_popup(self, flagged_df, source_label):
        if flagged_df is None or flagged_df.empty:
            return

        lines = self.build_flag_summary_lines(flagged_df)
        if not lines:
            return

        QMessageBox.information(
            self,
            "Flags Detected",
            "Flagged cause of\n\n"
            + "\n".join(lines)
            + "\n\nWhat happened to flagged rows:\n"
            + "- Title keyword flagged rows were set to INACTIVE and stock count was updated to 0.\n"
            + "- SKU phrase flagged rows were set to INACTIVE and stock count was updated to 0.\n"
            + "- Size flagged rows were set to INACTIVE and stock count was updated to 0."
        )

    def detect_account_from_filename(self, path):
        filename = Path(path).name.lower()
        if "84f77" in filename:
            return "Prabhu"
        if "946b8" in filename:
            return "Seema"
        return "Unknown"

    # ---------------- LOAD ----------------
    def load_file(self):
        self.load_size_overrides()
        self.load_flag_config()

        path, _ = QFileDialog.getOpenFileName(
            self, "Open", "", "Excel (*.xlsx *.xls *.csv)"
        )
        if not path:
            return

        if path.endswith(".csv"):
            self.df = pd.read_csv(path, encoding="latin1")
        elif path.endswith(".xlsx"):
            self.df = pd.read_excel(path)
        else:
            try:
                self.df = pd.read_excel(path)
            except Exception:
                import xlrd

                book = xlrd.open_workbook(path)
                sheet = book.sheet_by_index(0)
                data = [sheet.row_values(i) for i in range(sheet.nrows)]
                self.df = pd.DataFrame(data[1:], columns=data[0])

        self.file_type = path.split(".")[-1].lower()
        account_name = self.detect_account_from_filename(path)

        self.df.columns = self.df.columns.str.strip()
        self.df.reset_index(drop=True, inplace=True)
        self.df["__orig_index"] = self.df.index
        self.df["__locked"] = False
        self.df["__flag_exemptions"] = ""
        self.account_label.setText(f"Account: {account_name}")
        self.undo_state = None
        self.undo_label = ""
        self.change_log_entries = []
        self.change_log.clear()
        self.undo_btn.setEnabled(False)
        self.undo_btn.setText("Undo Last Move")

        self.available_modes = {
            "offer": all(col in self.df.columns for col in ["SKU ID", "FSN", "Selling Price(Rs)"]),
            "normal": all(col in self.df.columns for col in ["Seller SKU Id", "Product Title", "Bank Settlement"]),
        }

        if not any(self.available_modes.values()):
            QMessageBox.critical(self, "Error", "Unknown file format")
            return

        if self.available_modes["offer"] and not self.available_modes["normal"]:
            self.offer_mode = True
        elif self.available_modes["normal"]:
            self.offer_mode = False

        self.apply_mode_config()
        self.update_dashboard()

    # ---------------- CLEAN ----------------
    def clean_numeric(self, col):
        self.df[col] = (
            self.df[col].astype(str)
            .str.replace(",", "")
            .str.replace("₹", "")
            .str.replace("Rs", "")
            .str.strip()
        )
        self.df[col] = pd.to_numeric(self.df[col], errors="coerce")

    # ---------------- CLASSIFY ----------------
    def apply_classification(self):
        owners = ["Starvielle", "Genz Vane", "INDIVANE", "FADEVIELLE", "FLEECRANE"]

        def listing(title):
            return "Owner" if any(k.lower() in str(title).lower() for k in owners) else "Latched"

        def jeans(row):
            sku = str(row[self.col_sku]).lower()
            title = str(row[self.col_title]).lower()

            if "white" in sku:
                return "WHITE"
            if "ice" in sku or "blue" in sku:
                return "ICE"
            if "beige" in sku or "cream" in sku:
                return "BEIGE"
            if "baggy" in sku:
                return "BLACK-BAGGY"
            if "black" in sku and "relaxed" in title:
                return "BLACK-BAGGY"
            if "black" in sku:
                return "BLACK-PLAIN"
            return "MIX"

        self.df["Listing Type"] = self.df[self.col_title].apply(listing)
        self.df["Jeans Type"] = self.df.apply(jeans, axis=1)
        self.df["Size"] = self.df[self.col_sku].apply(self.detect_size_from_sku)
        if "Listing Status" in self.df.columns:
            self.df["Listing Status"] = self.df["Listing Status"].astype(str).str.strip().str.upper()
        flagged_df = self.apply_size_flags()
        self.sync_stock_count_with_status(pd.Series(True, index=self.df.index, dtype=bool))
        self.show_flag_popup(flagged_df, "size detection")

    # ---------------- OFFER LOGIC ----------------
    def apply_offer_logic(self, df):
        try:
            y = float(self.y_input.text()) / 100
            x = float(self.x_input.text()) / 100
            cap = float(self.cap_input.text())
        except:
            return df

        def compute(row):
            val = row[self.col_settlement]

            base = y * val
            discount = min(x * base, cap)
            final_price = val - discount

            j = row["Jeans Type"]

            try:
                threshold = float(self.thresholds[j].text())
            except:
                threshold = 0

            decision = "ACCEPT" if final_price >= threshold else "REJECT"

            return pd.Series([discount, final_price, decision])

        df[["Discount", "Final Price", "Decision"]] = df.apply(compute, axis=1)
        return df

    def compute_discount_only(self):
        try:
            y = float(self.y_input.text()) / 100
            x = float(self.x_input.text()) / 100
            cap = float(self.cap_input.text())
        except:
            QMessageBox.warning(self, "Error", "Invalid inputs")
            return

        def compute(row):
            val = row[self.col_settlement]

            base = y * val
            discount = min(x * base, cap)  # ✅ FIXED
            final_price = val - discount

            return pd.Series([discount, final_price])

        mask = self.get_filtered_mask(include_selection=True)
        if mask.empty or not mask.any():
            QMessageBox.warning(self, "Error", "No visible unlocked rows selected")
            return

        self.save_undo_state("Compute Discount")
        before_values = self.df.loc[mask, self.col_settlement].copy()
        self.df.loc[mask, ["Discount", "Final Price"]] = self.df.loc[mask].apply(compute, axis=1)
        after_values = self.df.loc[mask, self.col_settlement].copy()
        self.log_change("Compute Discount", before_values, after_values, int(mask.sum()), context_df=self.df.loc[mask])

        self.update_dashboard()
    
    def apply_decision_only(self):
        if "Final Price" not in self.df.columns:
            QMessageBox.warning(self, "Error", "Compute discount first")
            return

        def decide(row):
            j = row["Jeans Type"]

            try:
                threshold = float(self.thresholds[j].text())
            except:
                threshold = 0

            return "ACCEPT" if row["Final Price"] >= threshold else "REJECT"

        mask = self.get_filtered_mask(include_selection=True)
        if mask.empty or not mask.any():
            QMessageBox.warning(self, "Error", "No visible unlocked rows selected")
            return

        self.save_undo_state("Apply Decision")
        before_values = self.df.loc[mask, self.col_settlement].copy()
        self.df.loc[mask, "Decision"] = self.df.loc[mask].apply(decide, axis=1)
        after_values = self.df.loc[mask, self.col_settlement].copy()
        self.log_change("Apply Decision", before_values, after_values, int(mask.sum()), context_df=self.df.loc[mask])

        self.update_dashboard()

    def save_size_override(self):
        if self.df.empty:
            QMessageBox.warning(self, "Error", "Load a file first")
            return

        sku = self.size_sku_input.text().strip()
        if not sku:
            QMessageBox.warning(self, "Error", "Enter a SKU")
            return

        size = self.size_override_combo.currentText()
        self.save_undo_state("Save Size")
        self.size_overrides[sku] = size
        self.persist_size_overrides()

        mask = self.get_filtered_mask(include_selection=True)
        mask &= self.df[self.col_sku].astype(str).str.strip() == sku
        if mask.any():
            before_values = self.df.loc[mask, self.col_settlement].copy()
            self.df.loc[mask, "Size"] = size
            flagged_df = self.apply_size_flags(mask)
            self.sync_stock_count_with_status(mask)
            after_values = self.df.loc[mask, self.col_settlement].copy()
            self.log_change("Save Size", before_values, after_values, int(mask.sum()), f"SKU: {sku} -> {size}", self.df.loc[mask])
            self.show_flag_popup(flagged_df, "manual size override")
        else:
            QMessageBox.information(
                self,
                "Saved",
                "Size override saved for future lookups. Matching SKU not present in current selection."
            )

        self.size_sku_input.clear()
        self.setup_filters()
        self.update_dashboard()

    def download_undetected_sizes(self):
        if self.df.empty:
            QMessageBox.warning(self, "Error", "Load a file first")
            return

        mask = self.get_filtered_mask(include_selection=False)
        mask &= self.df["Size"].astype(str) == "UNDETECTED"
        export_df = self.df.loc[mask, [self.col_sku]].copy()
        export_df[self.col_sku] = export_df[self.col_sku].astype(str).str.strip()
        export_df = export_df.drop_duplicates(subset=[self.col_sku])
        export_df = export_df.rename(columns={self.col_sku: "sku"})
        export_df["size"] = ""

        if export_df.empty:
            QMessageBox.information(self, "Info", "No undetected size SKUs in the current view")
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Undetected Size Template",
            "undetected_sizes.xlsx",
            "Excel Files (*.xlsx)"
        )
        if not path:
            return

        if not path.endswith(".xlsx"):
            path += ".xlsx"

        export_df.to_excel(path, index=False)
        QMessageBox.information(self, "Success", f"Template saved with {len(export_df)} SKU rows")

    def upload_size_sheet(self):
        if self.df.empty:
            QMessageBox.warning(self, "Error", "Load a file first")
            return

        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Size Sheet",
            "",
            "Excel (*.xlsx *.xls *.csv)"
        )
        if not path:
            return

        if path.endswith(".csv"):
            upload_df = pd.read_csv(path, encoding="latin1")
        else:
            upload_df = pd.read_excel(path)

        upload_df.columns = [str(col).strip().lower() for col in upload_df.columns]
        if "sku" not in upload_df.columns or "size" not in upload_df.columns:
            QMessageBox.warning(self, "Error", "Upload file must contain 'sku' and 'size' columns")
            return

        upload_df = upload_df[["sku", "size"]].copy()
        upload_df["sku"] = upload_df["sku"].astype(str).str.strip()
        upload_df["size"] = upload_df["size"].astype(str).str.strip()
        upload_df = upload_df[
            (upload_df["sku"] != "") &
            (upload_df["size"].isin(self.size_values))
        ].drop_duplicates(subset=["sku"], keep="last")

        if upload_df.empty:
            QMessageBox.warning(self, "Error", "No valid sku/size rows found in upload")
            return

        self.save_undo_state("Upload Sizes")
        override_map = dict(zip(upload_df["sku"], upload_df["size"]))
        self.size_overrides.update(override_map)
        self.persist_size_overrides()

        current_sku = self.df[self.col_sku].astype(str).str.strip()
        mask = current_sku.isin(override_map.keys())
        before_values = self.df.loc[mask, self.col_settlement].copy()
        flagged_df = self.df.iloc[0:0].copy()
        if mask.any():
            self.df.loc[mask, "Size"] = current_sku[mask].map(override_map)
            flagged_df = self.apply_size_flags(mask)
            self.sync_stock_count_with_status(mask)
        after_values = self.df.loc[mask, self.col_settlement].copy()

        extra = f"Uploaded Size Rows: {len(upload_df)}"
        self.log_change("Upload Sizes", before_values, after_values, int(mask.sum()), extra, self.df.loc[mask])
        self.setup_filters()
        self.update_dashboard()
        self.show_flag_popup(flagged_df, "uploaded size sheet")
        QMessageBox.information(self, "Success", f"Imported {len(upload_df)} size mappings")

    # ---------------- FREEZE ----------------
    def apply_freeze(self):
        if self.df.empty:
            return

        col = self.freeze_col.currentText()
        val = self.freeze_val.text().lower()
        if not col or not val:
            return

        mask = self.get_filtered_mask(include_selection=True)
        mask &= self.df[col].astype(str).str.lower().str.contains(val, na=False)
        if not mask.any():
            return

        self.save_undo_state("Freeze")
        before_values = self.df.loc[mask, self.col_settlement].copy()
        self.df.loc[mask, "__locked"] = True
        after_values = self.df.loc[mask, self.col_settlement].copy()
        self.log_change("Freeze", before_values, after_values, int(mask.sum()), f"Column contains: {val}", self.df.loc[mask])
        self.update_dashboard()

    def unfreeze_all(self):
        if "__locked" in self.df.columns:
            self.save_undo_state("Unfreeze All")
            mask = self.df["__locked"] == True
            before_values = self.df.loc[mask, self.col_settlement].copy()
            self.df["__locked"] = False
            after_values = self.df.loc[mask, self.col_settlement].copy()
            self.log_change("Unfreeze All", before_values, after_values, int(mask.sum()), context_df=self.df.loc[mask])
        self.update_dashboard()

    # ---------------- BULK EDIT ----------------
    def apply_bulk_edit(self):
        if self.df.empty:
            return

        try:
            val = float(self.edit_value.text())
        except Exception:
            return

        cap_value = None
        cap_text = self.edit_cap_value.text().strip()
        if cap_text:
            try:
                cap_value = float(cap_text)
            except Exception:
                QMessageBox.warning(self, "Error", "Invalid cap value")
                return

        mask = self.get_filtered_mask(include_selection=True)
        if mask.empty or not mask.any():
            QMessageBox.warning(self, "Error", "No visible unlocked rows selected")
            return

        self.save_undo_state("Bulk Edit")
        before_values = self.df.loc[mask, self.col_settlement].copy()
        mode_name = self.edit_mode.currentText()
        if mode_name == "Add":
            self.df.loc[mask, self.col_settlement] += val
        elif mode_name == "Multiply":
            result = self.df.loc[mask, self.col_settlement] * val
            if cap_value is not None:
                if self.edit_cap_mode.currentText() == "Min":
                    result = np.minimum(result, cap_value)
                else:
                    result = np.maximum(result, cap_value)
            self.df.loc[mask, self.col_settlement] = result
        elif mode_name == "Replace":
            self.df.loc[mask, self.col_settlement] = val
        after_values = self.df.loc[mask, self.col_settlement].copy()
        extra = f"Mode: {mode_name}, Value: {val}"
        if mode_name == "Multiply" and cap_value is not None:
            extra += f", {self.edit_cap_mode.currentText()} Cap: {cap_value}"
        self.log_change("Bulk Edit", before_values, after_values, int(mask.sum()), extra, self.df.loc[mask])

        self.update_dashboard()

    def get_filtered_mask(self, include_selection=True, apply_status_filter=True):
        if self.df.empty:
            return pd.Series(False, index=self.df.index, dtype=bool)

        mask = (
            (self.df["__locked"] == False) &
            (self.df[self.col_settlement].notna()) &
            (self.df[self.col_settlement] <= self.slider.value())
        )

        search = self.search_box.text().lower()
        if search:
            mask &= (
                self.df[self.col_sku].astype(str).str.lower().str.contains(search, na=False) |
                self.df[self.col_title].astype(str).str.lower().str.contains(search, na=False)
            )

        if self.jeans_filter.currentText() != "All":
            mask &= self.df["Jeans Type"] == self.jeans_filter.currentText()

        if self.size_filter.currentText() != "All":
            mask &= self.df["Size"] == self.size_filter.currentText()

        if self.listing_filter.currentText() != "All":
            mask &= self.df["Listing Type"] == self.listing_filter.currentText()

        if apply_status_filter and "Listing Status" in self.df.columns and self.status_filter.currentText() != "All":
            mask &= self.df["Listing Status"] == self.status_filter.currentText()

        if include_selection and self.selected_range:
            low, high = self.selected_range
            mask &= (
                (self.df[self.col_settlement] >= low) &
                (self.df[self.col_settlement] <= high)
            )

        return mask

    def update_row_counts(self, visible_rows):
        loaded_rows = len(self.df)
        export_rows = len(self.df)
        self.row_count_label.setText(
            f"Loaded: {loaded_rows} | Visible: {visible_rows} | Export: {export_rows}"
        )

    def update_status_matrix(self, df):
        if df is None or df.empty or "Listing Status" not in df.columns or "Size" not in df.columns:
            self.status_matrix_label.setText("Size Status Matrix\nACTIVE: -\nINACTIVE: -")
            return

        status_series = df["Listing Status"].astype(str).str.strip().str.upper()
        size_series = df["Size"].astype(str).str.strip()

        active_sizes = sorted(size_series[status_series == "ACTIVE"].dropna().unique().tolist(), key=str)
        inactive_sizes = sorted(size_series[status_series == "INACTIVE"].dropna().unique().tolist(), key=str)

        active_text = ", ".join(active_sizes) if active_sizes else "-"
        inactive_text = ", ".join(inactive_sizes) if inactive_sizes else "-"
        self.status_matrix_label.setText(
            f"Size Status Matrix\nACTIVE: {active_text}\nINACTIVE: {inactive_text}"
        )

    def sync_stock_count_with_status(self, mask):
        stock_col = "Your Stock Count"
        status_col = "Listing Status"

        if stock_col not in self.df.columns or status_col not in self.df.columns:
            return

        if not pd.api.types.is_numeric_dtype(self.df[stock_col]):
            self.df[stock_col] = pd.to_numeric(self.df[stock_col], errors="coerce")

        active_mask = mask & (self.df[status_col].astype(str).str.strip().str.upper() == "ACTIVE")
        inactive_mask = mask & (self.df[status_col].astype(str).str.strip().str.upper() == "INACTIVE")

        if active_mask.any():
            self.df.loc[active_mask, stock_col] = 250
        if inactive_mask.any():
            self.df.loc[inactive_mask, stock_col] = 0

    def apply_status_change(self):
        if self.df.empty or "Listing Status" not in self.df.columns:
            QMessageBox.warning(self, "Error", "Listing Status column not available")
            return

        new_status = self.status_change_combo.currentText().strip().upper()
        mask = self.get_filtered_mask(include_selection=True)

        if mask.empty or not mask.any():
            QMessageBox.warning(self, "Error", "No visible unlocked rows to update")
            return

        self.save_undo_state("Set Status")
        if new_status == "ACTIVE":
            mask = self.resolve_activation_mask(mask, "Set Status")
            if mask.empty or not mask.any():
                QMessageBox.information(self, "No Changes", "Flagged rows stayed INACTIVE because the flags were not excluded.")
                self.update_dashboard()
                return

        before_values = self.df.loc[mask, self.col_settlement].copy()
        self.df.loc[mask, "Listing Status"] = new_status
        if new_status == "ACTIVE":
            self.apply_size_flags(mask)
        self.sync_stock_count_with_status(mask)
        after_values = self.df.loc[mask, self.col_settlement].copy()
        self.log_change("Set Status", before_values, after_values, int(mask.sum()), f"Status: {new_status}", self.df.loc[mask])
        self.setup_filters()
        self.update_dashboard()
    # ---------------- FILTER ----------------
    def setup_filters(self):
        current_listing = self.listing_filter.currentText()
        current_jeans = self.jeans_filter.currentText()
        current_size = self.size_filter.currentText()
        current_status = self.status_filter.currentText()

        self.listing_filter.clear()
        self.listing_filter.addItem("All")
        self.listing_filter.addItems(self.df["Listing Type"].dropna().astype(str).unique())

        self.jeans_filter.clear()
        self.jeans_filter.addItem("All")
        self.jeans_filter.addItems(self.df["Jeans Type"].dropna().astype(str).unique())

        self.size_filter.clear()
        self.size_filter.addItem("All")
        size_options = [size for size in self.size_values if size in self.df["Size"].astype(str).unique()]
        self.size_filter.addItems(size_options)
        if "UNDETECTED" in self.df["Size"].astype(str).unique():
            self.size_filter.addItem("UNDETECTED")

        self.status_filter.clear()
        self.status_filter.addItem("All")
        if "Listing Status" in self.df.columns:
            statuses = self.df["Listing Status"].dropna().astype(str).str.strip().str.upper().unique().tolist()
            statuses = [status for status in ["ACTIVE", "INACTIVE"] if status in statuses]
            self.status_filter.addItems(statuses)

        self.freeze_col.clear()
        self.freeze_col.addItems(self.df.columns)

        max_val = int(self.df[self.col_settlement].max() or 0)
        self.slider.setMaximum(max_val)
        self.slider.setValue(max_val)

        for combo, value in [
            (self.listing_filter, current_listing),
            (self.jeans_filter, current_jeans),
            (self.size_filter, current_size),
            (self.status_filter, current_status),
        ]:
            idx = combo.findText(value)
            if value and idx >= 0:
                combo.setCurrentIndex(idx)

    # ---------------- UPDATE ----------------
    def update_dashboard(self):
        if self.df.empty:
            self.update_row_counts(0)
            self.update_status_matrix(None)
            return

        chart_df = self.df[self.get_filtered_mask(include_selection=False)].copy()
        status_matrix_df = self.df[self.get_filtered_mask(include_selection=True, apply_status_filter=False)].copy()
        df = chart_df.copy()
        if self.selected_range:
            low, high = self.selected_range
            df = df[
                (df[self.col_settlement] >= low) &
                (df[self.col_settlement] <= high)
            ]

        self.update_row_counts(len(df))
        self.update_status_matrix(status_matrix_df)
        self.update_table(df)
        self.update_chart(chart_df)

    # ---------------- TABLE ----------------
    def update_table(self, df):
        self.table.blockSignals(True)
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns)

        for i in range(len(df)):
            for j in range(len(df.columns)):
                self.table.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))
        self.table.blockSignals(False)

    # ---------------- EDIT ----------------
    def handle_edit(self, item):
        row = item.row()
        col = item.column()
        header_item = self.table.horizontalHeaderItem(col)
        if header_item is None:
            return

        col_name = header_item.text()

        try:
            orig_idx_col = self.df.columns.get_loc("__orig_index")
            orig_idx = int(self.table.item(row, orig_idx_col).text())
            locked = self.df.loc[self.df["__orig_index"] == orig_idx, "__locked"].values[0]
            if locked:
                return

            row_mask = self.df["__orig_index"] == orig_idx
            before_values = self.df.loc[row_mask, self.col_settlement].copy()
            before_cell_value = str(self.df.loc[row_mask, col_name].iloc[0])
            self.save_undo_state(f"Edit {col_name}")
            self.df.loc[row_mask, col_name] = item.text()
            if col_name == self.col_sku:
                self.df.loc[row_mask, "Size"] = self.df.loc[row_mask, self.col_sku].apply(self.detect_size_from_sku)
                flagged_df = self.apply_size_flags(row_mask)
                self.sync_stock_count_with_status(row_mask)
                self.show_flag_popup(flagged_df, "SKU edit")
            elif col_name == "Listing Status":
                self.df.loc[row_mask, "Listing Status"] = str(item.text()).strip().upper()
                if str(item.text()).strip().upper() == "ACTIVE":
                    active_mask = self.resolve_activation_mask(row_mask, "Listing Status edit")
                    if active_mask.any():
                        self.df.loc[active_mask, "Listing Status"] = "ACTIVE"
                        self.apply_size_flags(active_mask)
                        self.sync_stock_count_with_status(active_mask)
                    inactive_rows = row_mask & ~active_mask
                    if inactive_rows.any():
                        self.df.loc[inactive_rows, "Listing Status"] = "INACTIVE"
                        self.sync_stock_count_with_status(inactive_rows)
                else:
                    self.df.loc[row_mask, "Listing Status"] = str(item.text()).strip().upper()
                    self.sync_stock_count_with_status(row_mask)
            after_values = self.df.loc[row_mask, self.col_settlement].copy()
            after_cell_value = str(self.df.loc[row_mask, col_name].iloc[0])
            extra = f"Row: {orig_idx}"
            if col_name == self.col_settlement:
                extra += f", Settlement: {before_cell_value} -> {after_cell_value}"
            elif col_name == "Listing Status":
                stock_value = str(self.df.loc[row_mask, "Your Stock Count"].iloc[0]) if "Your Stock Count" in self.df.columns else "unchanged"
                extra += f", Status: {before_cell_value} -> {after_cell_value}, Stock: {stock_value}"
            else:
                extra += f", {col_name}: {before_cell_value} -> {after_cell_value}"
            self.log_change(f"Edit {col_name}", before_values, after_values, 1, extra, self.df.loc[row_mask])
        except Exception:
            pass

    # ---------------- CHART ----------------
    def update_chart(self, df):
        for cid_name in ["hover_cid", "press_cid", "release_cid"]:
            cid = getattr(self, cid_name)
            if cid is not None:
                self.canvas.mpl_disconnect(cid)
                setattr(self, cid_name, None)

        self.figure.clear()
        ax = self.figure.add_subplot(111)

        settlement = df[self.col_settlement].dropna()
        if settlement.empty:
            self.canvas.draw()
            return

        freq = settlement.value_counts().sort_index()
        x = freq.index.tolist()
        y = freq.values.tolist()
        selected_mask = [False] * len(x)
        if self.selected_range:
            low, high = self.selected_range
            selected_mask = [low <= value <= high for value in x]

        colors = ["#d9480f" if is_selected else "#4c78a8" for is_selected in selected_mask]
        line = ax.plot(x, y, marker="o", color="#9ecae1", linewidth=2, zorder=2)[0]
        ax.scatter(x, y, c=colors, s=55, edgecolors="white", linewidths=0.8, zorder=3)

        selection_patch = None
        if self.selected_range:
            low, high = self.selected_range
            selection_patch = ax.axvspan(low, high, color="#f97316", alpha=0.16, zorder=1)

            in_range_x = [xv for xv, is_selected in zip(x, selected_mask) if is_selected]
            in_range_y = [yv for yv, is_selected in zip(y, selected_mask) if is_selected]
            if in_range_x:
                ax.plot(in_range_x, in_range_y, color="#d9480f", linewidth=2.6, zorder=4)

            padding = max((high - low) * 0.08, 1)
            ax.set_xlim(low - padding, high + padding)

        ax.set_facecolor("#fcfcfd")
        ax.grid(True, axis="y", alpha=0.18)
        ax.set_xlabel(f"Settlement values from Excel column: {self.col_settlement}")
        ax.set_ylabel("Frequency")
        ax.set_title(f"Settlement distribution from column: {self.col_settlement}")

        annot = ax.annotate(
            "",
            xy=(0, 0),
            xytext=(10, 10),
            textcoords="offset points",
            bbox=dict(boxstyle="round")
        )
        annot.set_visible(False)
        drag_patch = None

        def hover(event):
            nonlocal drag_patch

            if event.inaxes == ax and self.drag_start is not None and event.xdata is not None:
                low, high = sorted([self.drag_start, event.xdata])
                if drag_patch is not None:
                    drag_patch.remove()
                drag_patch = ax.axvspan(low, high, color="#fb923c", alpha=0.22, zorder=5)
                self.canvas.draw_idle()
                return

            if event.inaxes == ax:
                cont, ind = line.contains(event)
                if cont:
                    idx = ind["ind"][0]
                    annot.xy = (x[idx], y[idx])
                    annot.set_text(f"{x[idx]} : {y[idx]}")
                    annot.set_visible(True)
                    self.canvas.draw_idle()
                elif annot.get_visible():
                    annot.set_visible(False)
                    self.canvas.draw_idle()

        self.hover_cid = self.canvas.mpl_connect("motion_notify_event", hover)
        self.drag_start = None

        def press(event):
            nonlocal drag_patch

            if event.inaxes == ax and event.xdata is not None:
                self.drag_start = event.xdata
                if drag_patch is not None:
                    drag_patch.remove()
                    drag_patch = None

        def release(event):
            nonlocal drag_patch

            if event.inaxes != ax or self.drag_start is None or event.xdata is None:
                self.drag_start = None
                if drag_patch is not None:
                    drag_patch.remove()
                    drag_patch = None
                    self.canvas.draw_idle()
                return

            low, high = sorted([self.drag_start, event.xdata])
            if abs(high - low) < 0.5:
                self.drag_start = None
                if drag_patch is not None:
                    drag_patch.remove()
                    drag_patch = None
                    self.canvas.draw_idle()
                return

            before_values = df[self.col_settlement].copy()
            self.selected_range = (low, high)
            subset_df = df[
                (df[self.col_settlement] >= low) &
                (df[self.col_settlement] <= high)
            ].copy()
            subset = subset_df[self.col_settlement]

            if len(subset) > 0:
                try:
                    mode_value = mode(subset)
                except StatisticsError:
                    mode_value = "NA"

                self.stats_label.setText(
                    f"Count:{len(subset)} Mean:{np.mean(subset):.2f} "
                    f"Median:{np.median(subset):.2f} Mode:{mode_value}"
                )
                self.log_change(
                    "Range Select",
                    before_values,
                    subset,
                    len(subset_df),
                    f"Selected Range: {low:.2f} -> {high:.2f}",
                    subset_df
                )

            self.drag_start = None
            if drag_patch is not None:
                drag_patch.remove()
                drag_patch = None
            self.update_dashboard()

        self.press_cid = self.canvas.mpl_connect("button_press_event", press)
        self.release_cid = self.canvas.mpl_connect("button_release_event", release)

        self.canvas.draw()

    def reset_selection(self):
        self.selected_range = None
        self.stats_label.setText("No Selection")
        self.update_dashboard()

    # ---------------- EXPORT ----------------
    def export_file(self):
        if self.df.empty:
            QMessageBox.warning(self, "Error", "No data loaded")
            return

        df = self.df.sort_values("__orig_index")
        df = df.drop(columns=["__orig_index", "__locked", "__flag_exemptions"], errors="ignore")
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save File",
            f"PROGRAM OUTPUTTED.{self.file_type}",
            "All Files (*.*)"
        )

        if not path:
            return

        try:
            if self.file_type == "csv":
                df.to_csv(path, index=False)
            elif self.file_type == "xlsx":
                df.to_excel(path, index=False)
            elif self.file_type == "xls":
                if not path.endswith(".xlsx"):
                    path = path + ".xlsx"
                df.to_excel(path, index=False)
            else:
                df.to_excel(path + ".xlsx", index=False)

            QMessageBox.information(self, "Success", "File exported successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = Dashboard()
    w.show()
    sys.exit(app.exec_())
