import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QScrollArea, QLabel, QPushButton, QFrame, QSizePolicy, QLineEdit, QLabel
)
from PySide6.QtGui import QPixmap, QCursor, QDoubleValidator
from PySide6.QtCore import Qt

from docx import Document
from PIL import Image
import io

from utilities import require_file


DOCX_PATH = r"P:\ENGINEERING\Design Checklist\supporting_documents\reference_tab_info.docx"
EXCEL_PATH = r"P:\ENGINEERING\Design Checklist\supporting_documents\Flex Die Gap Calculator.xlsx"
ROTO_DIE_PATH = r"P:\CFS Documents\Roto Die List.xls"

# ---- Cache for parsed reference docx ----
_DOCX_CACHE = {"path": None, "mtime": None, "blocks": None}

def _get_docx_mtime(path):
    try:
        return os.path.getmtime(path)
    except Exception:
        return None

def get_reference_blocks(docx_path=DOCX_PATH):
    """
    Return parsed 'blocks' for the reference_tab_info.docx,
    reusing a cached parse unless the file changed on disk.
    """
    mtime = _get_docx_mtime(docx_path)
    if (_DOCX_CACHE["path"] != docx_path or
        _DOCX_CACHE["mtime"] != mtime or
        _DOCX_CACHE["blocks"] is None):
        # IMPORTANT: parse the file here (do NOT call get_reference_blocks again)
        blocks = parse_docx_blocks(docx_path)
        _DOCX_CACHE.update({"path": docx_path, "mtime": mtime, "blocks": blocks})
    return _DOCX_CACHE["blocks"]


def invalidate_reference_blocks_cache():
    """Call this if you need to force a rebuild (e.g., a manual 'Refresh')."""
    _DOCX_CACHE["blocks"] = None


class ScaledImageLabel(QLabel):
    def __init__(self, qpixmap, parent=None):
        super().__init__(parent)
        self.original_pixmap = qpixmap
        self.setScaledContents(False)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.original_pixmap:
            target_width = min(self.width(), self.original_pixmap.width())
            if target_width < 20:
                target_width = self.original_pixmap.width()
            ratio = self.original_pixmap.height() / self.original_pixmap.width()
            target_height = int(target_width * ratio)
            scaled = self.original_pixmap.scaled(
                target_width,
                target_height,
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
            self.setPixmap(scaled)

class ReferenceCalcWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QLineEdit, QLabel {
                font-size: 17px;
                min-height: 28px;
            }
        """)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(12)

        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("Enter value")
        self.input_edit.setFixedWidth(120)
        self.input_edit.setValidator(QDoubleValidator())

        eq_label = QLabel("→")

        self.result_edit = QLineEdit()
        self.result_edit.setReadOnly(True)
        self.result_edit.setPlaceholderText("Result")
        self.result_edit.setFixedWidth(120)

        layout.addWidget(QLabel("<b>Linear Inches:</b>"))
        layout.addWidget(self.input_edit)
        layout.addWidget(eq_label)
        layout.addWidget(QLabel("<b>Tonnage:</b>"))
        layout.addWidget(self.result_edit)
        layout.addStretch()

        self.input_edit.textChanged.connect(self.update_result)

    def clear_fields(self):
        self.input_edit.clear()
        self.result_edit.clear()
    
    def update_result(self, text):
        try:
            val = float(text)
            res = (val * 500) / 2000
            self.result_edit.setText(f"{res:.3f}")
        except Exception:
            self.result_edit.clear()

def create_reference_tab(tab_widget: QWidget):
    scroll_area = QScrollArea()
    scroll_area.setWidgetResizable(True)
    outer_widget = QWidget()
    layout = QVBoxLayout()
    layout.setContentsMargins(2, 2, 2, 2)
    layout.setSpacing(6)
    outer_widget.setLayout(layout)
    scroll_area.setWidget(outer_widget)

    # Collapsible sections (use cached blocks)
    blocks = get_reference_blocks(DOCX_PATH)
    build_collapsibles(layout, blocks)

    layout.addStretch()
    tab_layout = QVBoxLayout()
    tab_layout.setContentsMargins(0, 0, 0, 0)
    tab_layout.setSpacing(0)
    tab_layout.addWidget(scroll_area)
    tab_widget.setLayout(tab_layout)

def open_excel_calculator():
    if os.path.exists(EXCEL_PATH):
        os.startfile(EXCEL_PATH)
    else:
        from PySide6.QtWidgets import QMessageBox
        QMessageBox.critical(None, "File Not Found", f"Could not find file:\n{EXCEL_PATH}")

def open_roto_die_list():
    if os.path.exists(ROTO_DIE_PATH):
        os.startfile(ROTO_DIE_PATH)
    else:
        from PySide6.QtWidgets import QMessageBox
        QMessageBox.critical(None, "File Not Found", f"Could not find file:\n{ROTO_DIE_PATH}")

def parse_docx_blocks(path):
    from docx.text.paragraph import Paragraph
    from docx.table import Table

    # If the reference DOCX is missing, just return no blocks.
    # No popup, no crash.
    if not os.path.exists(path):
        return []

    doc = Document(path)
    blocks = []
    current_header = None
    current_content = []
    rels = doc.part._rels
    for element in doc.element.body.iterchildren():
        tag = element.tag.lower()
        if tag.endswith("}p"):
            para = Paragraph(element, doc)
            text = para.text.strip()
            style = para.style.name if para.style else ""
            if style.startswith("Heading") and text:
                if current_header is not None:
                    blocks.append((current_header, current_content))
                current_header = text
                current_content = []
            else:
                current_content.append(para)
        elif tag.endswith("}tbl"):
            tbl = Table(element, doc)
            current_content.append(tbl)
    if current_header is not None:
        blocks.append((current_header, current_content))
    return blocks


def build_collapsibles(parent_layout, blocks):
    for header, content in blocks:
        container = QWidget()
        container.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Maximum)
        vbox = QVBoxLayout(container)
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.setSpacing(4)

        # Header button: larger/bolder
        header_btn = QPushButton(f"➔ {header}")
        header_btn.setStyleSheet("""
            QPushButton {
                text-align: left;
                font-weight: bold;
                background: #f7f7f7;
                border: 1px solid #bbbbbb;
                border-radius: 2px;
                padding: 3px 8px;
                font-size: 13px;
                min-height: 21px;
            }
        """)
        header_btn.setCursor(QCursor(Qt.PointingHandCursor))
        header_btn.setFixedHeight(26)
        header_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        vbox.addWidget(header_btn)

        content_frame = QWidget()
        content_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        content_layout = QVBoxLayout(content_frame)
        content_layout.setContentsMargins(14, 0, 12, 7)
        content_layout.setSpacing(3)
        content_frame.setVisible(False)
        vbox.addWidget(content_frame)

        def make_toggle(frame, btn, text):
            def toggle():
                if frame.isVisible():
                    frame.setVisible(False)
                    btn.setText(f"➔ {text}")
                else:
                    frame.setVisible(True)
                    btn.setText(f"▼ {text}")
            return toggle
        header_btn.clicked.connect(make_toggle(content_frame, header_btn, header))

        for item_type, item in content:
            if item_type == "text":
                lbl = QLabel(item)
                lbl.setWordWrap(True)
                lbl.setStyleSheet("font-size: 13px; margin: 0; padding: 0;")
                lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                content_layout.addWidget(lbl)
            elif item_type == "table":
                render_table(content_layout, item)
            elif item_type == "image":
                render_image(content_layout, item)

        parent_layout.addWidget(container)

def render_table(parent_layout, table):
    from PySide6.QtWidgets import QGridLayout, QSizePolicy

    frame = QFrame()
    grid = QGridLayout()
    grid.setContentsMargins(2, 2, 2, 2)
    grid.setHorizontalSpacing(12)
    grid.setVerticalSpacing(3)
    grid.setSizeConstraint(QGridLayout.SetMinAndMaxSize)
    frame.setLayout(grid)
    frame.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Maximum)

    col_count = max(len(r.cells) for r in table.rows)
    font_metrics_normal = frame.fontMetrics()

    # Bold font metrics for header sizing
    bold_lbl = QLabel()
    bold_lbl.setStyleSheet("font-weight: bold;")
    font_metrics_bold = bold_lbl.fontMetrics()

    # First pass: measure max width for each column
    col_widths = [0] * col_count
    PAD = 14  # a little padding so text isn't flush
    for r, row in enumerate(table.rows):
        for c in range(len(row.cells)):
            text = row.cells[c].text.strip()
            fm = font_metrics_bold if r == 0 else font_metrics_normal
            width = fm.horizontalAdvance(text) + PAD
            if width > col_widths[c]:
                col_widths[c] = width

    # Apply widths and add labels
    for r, row in enumerate(table.rows):
        for c in range(len(row.cells)):
            text = row.cells[c].text.strip()
            lbl = QLabel(text)
            lbl.setWordWrap(False)  # don't wrap; rely on auto width
            if r == 0:
                lbl.setStyleSheet("font-weight: bold; font-size: 13px;")
            else:
                lbl.setStyleSheet("font-size: 13px;")
            if c == 0:
                lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            else:
                lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            grid.addWidget(lbl, r, c)
            grid.setColumnMinimumWidth(c, col_widths[c])
            grid.setColumnStretch(c, 0)

    parent_layout.addWidget(frame, 0, Qt.AlignLeft)

def render_image(parent_layout, stream):
    try:
        img = Image.open(io.BytesIO(stream))
        data = io.BytesIO()
        img.save(data, format='PNG')
        qimg = QPixmap()
        qimg.loadFromData(data.getvalue(), 'PNG')
        label = ScaledImageLabel(qimg)
        label.setMinimumHeight(30)
        label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        parent_layout.addWidget(label)
    except Exception:
        label = QLabel("[Image render failed]")
        parent_layout.addWidget(label)
