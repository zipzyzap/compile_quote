import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QCheckBox, QLineEdit,
    QLabel, QMessageBox, QDoubleSpinBox, QFrame, QSizePolicy, QScrollArea, QGridLayout
)
from PySide6.QtGui import QFontMetrics
from PySide6.QtCore import Qt

from utilities import (
    STRATASYS_ORDER, FORMLABS_HEADERS,
    calculate_stratasys_cost, calculate_stratasys_3d_cost,
    calculate_formlabs_cost, calculate_formlabs_3d_cost
)
import email_gen

EXCEL_PATH = r"P:\ENGINEERING\Design Checklist\supporting_documents\checklist_questions.xlsx"

def auto_resize_lineedit(lineedit, min_width=110, max_width=250):
    fm = lineedit.fontMetrics()
    text = lineedit.text() or lineedit.placeholderText()
    width = fm.horizontalAdvance(text) + 18
    lineedit.setFixedWidth(min(max(width, min_width), max_width))

class ActionBar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        row = QHBoxLayout(self)
        row.setContentsMargins(6, 2, 6, 2)
        row.setSpacing(10)

        self.btn_remove = QPushButton("Remove")
        self.btn_remove.setStyleSheet("""
            QPushButton {
                background: #ffeaea;
                color: #a10000;
                border: 1.2px solid #e0a4a4;
                border-radius: 7px;
                font-weight: 600;
                padding: 3px 15px;
            }
            QPushButton:hover {
                background: #fff0f0;
            }
        """)
        row.addWidget(self.btn_remove)

        self.btn_open = QPushButton("Open Drawing")
        row.addWidget(self.btn_open)

        self.btn_3d = QPushButton("3D Quote")
        self.btn_3d.setCheckable(True)
        self.btn_3d.setChecked(False)
        row.addWidget(self.btn_3d)

        self.btn_include = QPushButton("Include in Email")
        self.btn_include.setCheckable(True)
        self.btn_include.setChecked(True)
        row.addWidget(self.btn_include)

        row.addStretch()

class QuoteInfoTab(QWidget):
    def __init__(self, dirty_tracker=None, parent=None):
        super().__init__(parent)
        self.dirty_tracker = dirty_tracker
        self._loading = False
        self.rows = []
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(5)

        header = QWidget()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(6)
        btn_add = QPushButton("Add Line")
        btn_add.setFixedHeight(22)
        btn_add.setStyleSheet("font-weight: bold;")
        btn_add.clicked.connect(self.add_row)
        btn_email = QPushButton("Generate Email")
        btn_email.setFixedHeight(22)
        btn_email.setStyleSheet("font-weight: bold;")
        btn_email.clicked.connect(self.handle_generate_email)
        header_layout.addWidget(btn_add)
        header_layout.addWidget(btn_email)
        header_layout.addStretch()
        header.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        layout.addWidget(header)

        note_lbl = QLabel(
            "Do not need to include rev level in drawing number\n"
            "Only drawings that are marked will generate emails"
        )
        note_lbl.setStyleSheet("font-size: 11px; color: #555; margin: 2px 0 0 0;")
        note_lbl.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        layout.addWidget(note_lbl)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.row_container = QWidget()
        self.row_container_layout = QVBoxLayout(self.row_container)
        self.row_container_layout.setContentsMargins(0, 0, 0, 0)
        self.row_container_layout.setSpacing(1)
        self.row_container.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Maximum)
        self.scroll.setWidget(self.row_container)
        layout.addWidget(self.scroll)

        self.add_row()

    def add_row(self, initial_data=None):
        row = QuoteInfoRow(self, self.dirty_tracker, self, initial_data)
        self.rows.append(row)
        self.row_container_layout.addWidget(row)
        scroll = self.scroll.verticalScrollBar()
        scroll.setValue(scroll.maximum())

    def remove_row(self, row_widget):
        self.rows.remove(row_widget)
        row_widget.setParent(None)
        row_widget.deleteLater()
        if self.dirty_tracker and not self._loading:
            self.dirty_tracker.mark_dirty()

    def handle_generate_email(self):
        quote_data = self.get_quote_info_data()
        email_gen.generate_emails(quote_data, parent=self)

    def get_quote_info_data(self, include_all=False):
        result = []
        for row in self.rows:
            data = row.get_row_data()
            if include_all or data.get("include", True):
                result.append(data)
        return result

    def load_quote_info_data(self, data):
        self._loading = True
        self.clear_quote_info_tab(skip_add=True)
        for rowdata in data:
            self.add_row(initial_data=rowdata)
        self._loading = False

    def clear_quote_info_tab(self, skip_add=False):
        self._loading = True
        for row in self.rows:
            row.setParent(None)
            row.deleteLater()
        self.rows.clear()
        if not skip_add:
            self.add_row()
        self._loading = False

class QuoteInfoRow(QWidget):
    def __init__(self, tab: QWidget, dirty_tracker=None, parent=None, initial_data=None):
        super().__init__(parent)
        self.tab = tab
        self.dirty_tracker = dirty_tracker
        self._loading = False

        self.outer_frame = QFrame(self)
        self.outer_frame.setFrameShape(QFrame.StyledPanel)
        self.outer_frame.setStyleSheet("""
            QFrame {
                border: 2px solid #c4c4c4;
                border-radius: 8px;
                background: #fcfcfc;
            }
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 12)
        layout.addWidget(self.outer_frame)
        frame_layout = QVBoxLayout(self.outer_frame)
        frame_layout.setContentsMargins(8, 8, 8, 8)
        frame_layout.setSpacing(4)
        self.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)

        self.action_bar = ActionBar()
        frame_layout.addWidget(self.action_bar)
        self.action_bar.btn_remove.clicked.connect(lambda: self.tab.remove_row(self))
        self.action_bar.btn_open.clicked.connect(self.open_drawing)
        self.action_bar.btn_3d.toggled.connect(lambda checked: self._on_3d_toggled(checked))
        self.action_bar.btn_include.toggled.connect(lambda checked: self._on_include_toggled(checked))

        field_row = QWidget()
        field_row_layout = QHBoxLayout(field_row)
        field_row_layout.setContentsMargins(0, 0, 0, 0)
        field_row_layout.setSpacing(6)
        placeholders = ["Drawing Number", "Material", "Quantity", "Customer Part Number"]
        self.fields = []
        for ph in placeholders:
            field = QLineEdit()
            field.setPlaceholderText(ph)
            auto_resize_lineedit(field)
            field.textChanged.connect(lambda _, f=field: self.on_field_changed(f))
            field.setStyleSheet("font-size: 12px;")
            field_row_layout.addWidget(field)
            self.fields.append(field)
        frame_layout.addWidget(field_row)

        self.table_frame = QWidget()
        self.table_frame.setVisible(self.action_bar.btn_3d.isChecked())
        frame_layout.addWidget(self.table_frame)
        table_layout = QVBoxLayout(self.table_frame)
        table_layout.setContentsMargins(2, 2, 2, 2)
        table_layout.setSpacing(4)

        s_header = QLabel("Stratasys")
        s_header.setStyleSheet("font-weight: bold; font-size: 12px; border: none; background: transparent;")
        table_layout.addWidget(s_header)
        stratasys_grid = QGridLayout()
        stratasys_grid.setHorizontalSpacing(4)
        stratasys_grid.setVerticalSpacing(2)
        table_layout.addLayout(stratasys_grid)

        s_headers = list(STRATASYS_ORDER) + ["Time (hrs)", "Material $", "3D Cost"]
        self.s_vars = {}
        for col, label in enumerate(s_headers):
            header_lbl = QLabel(label)
            header_lbl.setAlignment(Qt.AlignCenter)
            header_lbl.setStyleSheet("font-size: 11px; border: none; background: transparent;")
            stratasys_grid.addWidget(header_lbl, 0, col)
        for col, mat in enumerate(STRATASYS_ORDER):
            le = QLineEdit()
            le.setFixedWidth(48)
            le.setStyleSheet("font-size: 12px;")
            le.textChanged.connect(lambda _, m=mat: self.on_s_var_changed())
            self.s_vars[mat] = le
            stratasys_grid.addWidget(le, 1, col)
        self.s_time = QLineEdit()
        self.s_time.setFixedWidth(48)
        self.s_time.setStyleSheet("font-size: 12px;")
        self.s_time.textChanged.connect(self.on_s_var_changed)
        stratasys_grid.addWidget(self.s_time, 1, len(STRATASYS_ORDER))
        self.s_material = QLineEdit()
        self.s_material.setReadOnly(True)
        self.s_material.setFixedWidth(70)
        stratasys_grid.addWidget(self.s_material, 1, len(STRATASYS_ORDER) + 1)
        self.s_3dcost = QLineEdit()
        self.s_3dcost.setReadOnly(True)
        self.s_3dcost.setFixedWidth(70)
        stratasys_grid.addWidget(self.s_3dcost, 1, len(STRATASYS_ORDER) + 2)

        f_header = QLabel("Formlabs")
        f_header.setStyleSheet("font-weight: bold; font-size: 12px; border: none; background: transparent;")
        table_layout.addWidget(f_header)
        formlabs_grid = QGridLayout()
        formlabs_grid.setHorizontalSpacing(4)
        formlabs_grid.setVerticalSpacing(2)
        table_layout.addLayout(formlabs_grid)
        f_headers = list(FORMLABS_HEADERS)
        self.f_vars = {}
        for col, label in enumerate(f_headers):
            header_lbl = QLabel(label)
            header_lbl.setAlignment(Qt.AlignCenter)
            header_lbl.setStyleSheet("font-size: 11px; border: none; background: transparent;")
            formlabs_grid.addWidget(header_lbl, 0, col)
        for col, h in enumerate(f_headers):
            le = QLineEdit()
            if h in ("Material $", "3D Cost"):
                le.setReadOnly(True)
                le.setFixedWidth(70)
            else:
                le.setFixedWidth(48)
                le.textChanged.connect(self.on_f_var_changed)
            le.setStyleSheet("font-size: 12px;")
            self.f_vars[h] = le
            formlabs_grid.addWidget(le, 1, col)

        if initial_data:
            self._loading = True
            for i, value in enumerate(initial_data.get("fields", [])):
                self.fields[i].setText(value)
            self.action_bar.btn_3d.setChecked(initial_data.get("enable_3d", False))
            self.action_bar.btn_include.setChecked(initial_data.get("include", True))
            for mat, val in initial_data.get("stratasys", {}).items():
                if mat in self.s_vars:
                    self.s_vars[mat].setText(val)
            self.s_time.setText(initial_data.get("stratasys", {}).get("Time (hrs)", ""))
            self.s_material.setText(initial_data.get("stratasys", {}).get("Material $", "$0.00"))
            self.s_3dcost.setText(initial_data.get("stratasys", {}).get("3D Cost", "$0.00"))
            for mat, val in initial_data.get("formlabs", {}).items():
                if mat in self.f_vars:
                    self.f_vars[mat].setText(val)
            self.update_costs()
            self._loading = False

    def on_field_changed(self, field):
        if self._loading or not self.dirty_tracker:
            return
        auto_resize_lineedit(field)
        self.dirty_tracker.mark_dirty()

    def on_s_var_changed(self):
        if self._loading:
            return
        self.update_costs()
        if self.dirty_tracker:
            self.dirty_tracker.mark_dirty()

    def on_f_var_changed(self):
        if self._loading:
            return
        self.update_costs()
        if self.dirty_tracker:
            self.dirty_tracker.mark_dirty()

    def _on_3d_toggled(self, checked):
        if self._loading:
            self.toggle_3d_tables(checked)
            return
        self.toggle_3d_tables(checked)
        if self.dirty_tracker:
            self.dirty_tracker.mark_dirty()

    def open_drawing(self):
        drawing_num = self.fields[0].text().strip()
        if not drawing_num or drawing_num == "Drawing Number":
            QMessageBox.warning(self, "No Drawing Number", "Please enter a drawing number first.")
            return
        try:
            from utilities import find_latest_pdf_with_rev
        except ImportError:
            def find_latest_pdf_with_rev(x): return (None, None)
        pdf_path, _ = find_latest_pdf_with_rev(drawing_num)
        import os
        if pdf_path and os.path.exists(pdf_path):
            os.startfile(pdf_path)
        else:
            QMessageBox.warning(self, "Drawing Not Found", f"No drawing found for: {drawing_num}")

    def toggle_3d_tables(self, checked):
        self.table_frame.setVisible(checked)

    def update_costs(self):
        if not self.action_bar.btn_3d.isChecked():
            return
        s = {}
        for k in STRATASYS_ORDER:
            val = self.s_vars[k].text()
            try:
                s[k] = float(val)
            except Exception:
                s[k] = 0.0
        try:
            t_val = float(self.s_time.text())
        except Exception:
            t_val = 0.0
        total = calculate_stratasys_cost(s)
        self.s_material.setText(f"${total:.2f}")
        self.s_3dcost.setText(f"${calculate_stratasys_3d_cost(total, t_val):.2f}")
        f = {}
        for k, v in self.f_vars.items():
            if isinstance(v, QLineEdit) and k not in ("Material $", "3D Cost"):
                try:
                    f[k] = float(v.text())
                except Exception:
                    f[k] = 0.0
        try:
            t_f = float(self.f_vars["Time (hrs)"].text())
        except Exception:
            t_f = 0.0
        mat = calculate_formlabs_cost(f)
        if "Material $" in self.f_vars:
            self.f_vars["Material $"].setText(f"${mat:.2f}")
        if "3D Cost" in self.f_vars:
            self.f_vars["3D Cost"].setText(f"${calculate_formlabs_3d_cost(mat, t_f):.2f}")

    def get_row_data(self):
        placeholders = ["Drawing Number", "Material", "Quantity", "Customer Part Number"]
        fields = [self.fields[i].text().strip() if self.fields[i].text().strip() != placeholders[i] else "" for i in range(4)]
        return {
            "include": self.action_bar.btn_include.isChecked(),
            "fields": fields,
            "enable_3d": self.action_bar.btn_3d.isChecked(),
            "stratasys": {k: self.s_vars[k].text() for k in STRATASYS_ORDER} | {
                "Time (hrs)": self.s_time.text(),
                "Material $": self.s_material.text(),
                "3D Cost": self.s_3dcost.text()
            },
            "formlabs": {k: self.f_vars[k].text() for k in FORMLABS_HEADERS}
        }
        
    def _on_include_toggled(self, checked):
        if self._loading:
            return
        if self.dirty_tracker:
            self.dirty_tracker.mark_dirty()

_tab_instance = None

def create_quote_info_tab(tab_widget: QWidget, dirty_tracker=None):
    global _tab_instance
    _tab_instance = QuoteInfoTab(dirty_tracker)
    layout = QVBoxLayout(tab_widget)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.addWidget(_tab_instance)

def get_quote_info_data(include_all=False):
    return _tab_instance.get_quote_info_data(include_all=include_all) if _tab_instance else []

def load_quote_info_data(data):
    if _tab_instance:
        _tab_instance.load_quote_info_data(data)

def clear_quote_info_tab(skip_add=False):
    if _tab_instance:
        _tab_instance.clear_quote_info_tab(skip_add=skip_add)
