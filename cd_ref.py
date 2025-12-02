import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QSizePolicy, QMessageBox
)
from PySide6.QtGui import QDoubleValidator, QFont, QColor
from PySide6.QtCore import Qt

tool_data = [
    {"width": 10, "tooth_count": 72},
    {"width": 10, "tooth_count": 84},
    {"width": 10, "tooth_count": 89},
    {"width": 10, "tooth_count": 90},
    {"width": 10, "tooth_count": 100},
    {"width": 10, "tooth_count": 113},
    {"width": 13, "tooth_count": 113},
    {"width": 10, "tooth_count": 118},
    {"width": 10, "tooth_count": 122},
    {"width": 10, "tooth_count": 138},
    {"width": 10, "tooth_count": 148},
    {"width": 13, "tooth_count": 158},
]

ROTO_DIE_PATH  = r"P:\CFS Documents\Roto Die List.xls"
PUNCH_DIE_PATH = r"P:\CFS Documents\Punch Die List.xlsx"  # <— NEW

MIN_CAVITY_GAP = 0.062
MAX_CAVITY_GAP = 2.000

def open_roto_die_list():
    if os.path.exists(ROTO_DIE_PATH):
        os.startfile(ROTO_DIE_PATH)
    else:
        QMessageBox.critical(None, "File Not Found", f"Could not find file:\n{ROTO_DIE_PATH}")

def open_punch_die_list():  # <— NEW
    if os.path.exists(PUNCH_DIE_PATH):
        os.startfile(PUNCH_DIE_PATH)
    else:
        QMessageBox.critical(None, "File Not Found", f"Could not find file:\n{PUNCH_DIE_PATH}")

class NumericTableWidgetItem(QTableWidgetItem):
    def __init__(self, text, value):
        super().__init__(text)
        self.setData(Qt.UserRole, value)
    def __lt__(self, other):
        if isinstance(other, QTableWidgetItem):
            return self.data(Qt.UserRole) < other.data(Qt.UserRole)
        return super().__lt__(other)

class GapCalculatorWidget(QWidget):  # keep name for compatibility
    def __init__(self):
        super().__init__()

        bold_label_font = QFont()
        bold_label_font.setPointSize(10)
        bold_label_font.setBold(True)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(4, 4, 4, 4)
        main_layout.setSpacing(4)
        main_layout.setAlignment(Qt.AlignTop)

        # --- Row 1: centered buttons ---
        buttons_row = QHBoxLayout()
        buttons_row.setContentsMargins(0, 0, 0, 0)
        buttons_row.setSpacing(8)

        btn_roto = QPushButton("Roto Die List")
        btn_roto.setFixedHeight(26)
        btn_roto.setStyleSheet("padding: 3px 10px; font-size: 10pt; font-weight: bold;")
        btn_roto.clicked.connect(open_roto_die_list)

        btn_punch = QPushButton("Punch Die List")
        btn_punch.setFixedHeight(26)
        btn_punch.setStyleSheet("padding: 3px 10px; font-size: 10pt; font-weight: bold;")
        btn_punch.clicked.connect(open_punch_die_list)

        buttons_row.addStretch()
        buttons_row.addWidget(btn_roto)
        buttons_row.addWidget(btn_punch)
        buttons_row.addStretch()
        main_layout.addLayout(buttons_row)

        # --- Row 2: centered "Linear Inches → Tonnage" ---
        li_row = QHBoxLayout()
        li_row.setContentsMargins(0, 0, 0, 0)
        li_row.setSpacing(6)

        li_label = QLabel("Linear Inches:")
        li_label.setFont(bold_label_font)

        self.tonnage_input = QLineEdit()
        self.tonnage_input.setPlaceholderText("Enter value")
        self.tonnage_input.setFixedWidth(140)
        self.tonnage_input.setValidator(QDoubleValidator(0.001, 99999.999, 3))

        arrow_lbl = QLabel("  →  ")

        result_label = QLabel("Tonnage:")
        result_label.setFont(bold_label_font)

        self.tonnage_result = QLineEdit()
        self.tonnage_result.setReadOnly(True)
        self.tonnage_result.setPlaceholderText("Result")
        self.tonnage_result.setFixedWidth(140)

        li_row.addStretch()
        li_row.addWidget(li_label)
        li_row.addWidget(self.tonnage_input)
        li_row.addWidget(arrow_lbl)
        li_row.addWidget(result_label)
        li_row.addWidget(self.tonnage_result)
        li_row.addStretch()
        main_layout.addLayout(li_row)

        # --- Row 3: left "Length" and right "Width" ---
        dims_row = QHBoxLayout()
        dims_row.setContentsMargins(0, 0, 0, 0)
        dims_row.setSpacing(12)

        # Left group
        left_group = QHBoxLayout()
        left_group.setContentsMargins(0, 0, 0, 0)
        left_group.setSpacing(6)
        length_label = QLabel("Web Direction Length:")
        length_label.setFont(bold_label_font)
        self.length_input = QLineEdit()
        self.length_input.setFixedWidth(140)
        self.length_input.setValidator(QDoubleValidator(0.001, 999.999, 3))
        left_group.addWidget(length_label)
        left_group.addWidget(self.length_input)

        # Right group
        right_group = QHBoxLayout()
        right_group.setContentsMargins(0, 0, 0, 0)
        right_group.setSpacing(6)
        width_label = QLabel("Web Direction Width:")
        width_label.setFont(bold_label_font)
        self.width_input = QLineEdit()
        self.width_input.setFixedWidth(140)
        self.width_input.setValidator(QDoubleValidator(0.001, 999.999, 3))
        right_group.addWidget(width_label)
        right_group.addWidget(self.width_input)

        dims_row.addLayout(left_group)
        dims_row.addStretch()
        dims_row.addLayout(right_group)

        main_layout.addLayout(dims_row)
        main_layout.addSpacing(6)


        # Side-by-side tables
        side_by_side = QHBoxLayout()
        side_by_side.setSpacing(20)

        self.table_length = self._build_table()
        self.table_width = self._build_table()

        self.length_container = QWidget()
        length_layout = QVBoxLayout(self.length_container)
        length_layout.setContentsMargins(0, 0, 0, 0)
        length_layout.setSpacing(0)
        length_layout.addWidget(self.table_length)

        self.width_container = QWidget()
        width_layout = QVBoxLayout(self.width_container)
        width_layout.setContentsMargins(0, 0, 0, 0)
        width_layout.setSpacing(0)
        width_layout.addWidget(self.table_width)

        side_by_side.addWidget(self.length_container)
        side_by_side.addWidget(self.width_container)

        main_layout.addLayout(side_by_side)

        # Signals
        self.tonnage_input.textChanged.connect(self.update_tonnage)
        self.length_input.textChanged.connect(self.update_tables)
        self.width_input.textChanged.connect(self.update_tables)

        self.length_input.editingFinished.connect(lambda: self.format_input_decimal(self.length_input))
        self.width_input.editingFinished.connect(lambda: self.format_input_decimal(self.width_input))
        self.tonnage_input.editingFinished.connect(lambda: self.format_input_decimal(self.tonnage_input))

        # Tab order
        self.setTabOrder(self.length_input, self.width_input)

    def _try_parse_float(self, text):
        try:
            return float(text)
        except ValueError:
            return None

    def _build_table(self):
        table = QTableWidget()
        table.setColumnCount(5)
        table.setHorizontalHeaderLabels([
            "Tooth Count", "Tool Width", "Cylinder Length", "Cavities", "Cavity Gap"
        ])
        table.verticalHeader().setVisible(False)
        table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        table.setFrameShape(QTableWidget.NoFrame)
        table.setShowGrid(True)
        table.setWordWrap(False)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionMode(QTableWidget.NoSelection)
        return table

    def format_input_decimal(self, line_edit):
        text = line_edit.text().strip()
        try:
            val = float(text)
            if not text.endswith(".000") and '.' not in text:
                line_edit.setText(f"{val:.3f}")
        except ValueError:
            line_edit.clear()

    def update_tables(self, _=None):
        length_text = self.length_input.text().strip()
        width_text = self.width_input.text().strip()

        length_val = self._try_parse_float(length_text)
        width_val = self._try_parse_float(width_text)

        if length_val is not None:
            data_length = self._calculate_gap_data(length_val, material_width=width_val)
            if data_length:
                self._populate_table(self.table_length, data_length)
                self.length_container.setMaximumHeight(16777215)
            else:
                self.table_length.clearContents()
                self.table_length.setRowCount(0)
                self.length_container.setMaximumHeight(0)
        else:
            self.table_length.clearContents()
            self.table_length.setRowCount(0)
            self.length_container.setMaximumHeight(0)

        if width_val is not None:
            data_width = self._calculate_gap_data(width_val, material_width=length_val)
            if data_width:
                self._populate_table(self.table_width, data_width)
                self.width_container.setMaximumHeight(16777215)
            else:
                self.table_width.clearContents()
                self.table_width.setRowCount(0)
                self.width_container.setMaximumHeight(0)
        else:
            self.table_width.clearContents()
            self.table_width.setRowCount(0)
            self.width_container.setMaximumHeight(0)

    def calculate_length_table(self):
        text = self.length_input.text().strip()
        try:
            length_val = float(text)
        except ValueError:
            self.table_length.clearContents()
            return
        self.length_input.setText(f"{length_val:.3f}")
        data = self._calculate_gap_data(length_val)
        self._populate_table(self.table_length, data)

    def calculate_width_table(self):
        text = self.width_input.text().strip()
        try:
            width_val = float(text)
        except ValueError:
            self.table_width.clearContents()
            return
        self.width_input.setText(f"{width_val:.3f}")
        data = self._calculate_gap_data(width_val, material_width=length_val)
        self._populate_table(self.table_width, data)

    def _calculate_gap_data(self, web_direction_value, material_width=None):
        results = []
        for tool in tool_data:
            width = tool["width"]
            tooth_count = tool["tooth_count"]
            cylinder_length = (tooth_count * 0.125)

            if material_width is not None and material_width > width:
                continue

            for cavities in range(1, 26):
                gap = (cylinder_length / cavities) - web_direction_value
                if MIN_CAVITY_GAP <= gap <= MAX_CAVITY_GAP:
                    results.append([
                        tooth_count,
                        f'{width}"',
                        f'{cylinder_length:.3f}"',
                        cavities,
                        gap
                    ])
        return results

    def _populate_table(self, table, data):
        table.setRowCount(len(data))
        for r, row in enumerate(data):
            for c, val in enumerate(row):
                if c == 4:
                    item = NumericTableWidgetItem(f"{val:.3f}", val)
                    gap = val
                    if gap <= 0.375:
                        item.setBackground(QColor(200, 255, 200))
                    elif gap <= 1.000:
                        item.setBackground(QColor(255, 255, 200))
                    else:
                        item.setBackground(QColor(255, 200, 200))
                else:
                    item = QTableWidgetItem(str(val))
                    item.setTextAlignment(Qt.AlignCenter)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                table.setItem(r, c, item)
        table.resizeColumnsToContents()
        table.resizeRowsToContents()
        table.sortItems(4, Qt.AscendingOrder)

    def clear_fields(self):
        self.tonnage_input.clear()
        self.tonnage_result.clear()
        self.length_input.clear()
        self.width_input.clear()

        self.table_length.clearContents()
        self.table_length.setRowCount(0)
        self.length_container.setMaximumHeight(0)

        self.table_width.clearContents()
        self.table_width.setRowCount(0)
        self.width_container.setMaximumHeight(0)

    def update_tonnage(self, text):
        try:
            val = float(text)
            res = (val * 500) / 2000
            self.tonnage_result.setText(f"{res:.3f}")
        except Exception:
            self.tonnage_result.clear()
