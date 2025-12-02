import os
import base64
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit,
    QTextEdit, QLabel, QFrame, QSizePolicy, QDialog, QScrollArea,
    QGridLayout
)
from PySide6.QtGui import QPixmap, QImage, QGuiApplication, QPalette, QColor
from PySide6.QtCore import Qt, QBuffer, QTimer, QEvent

THUMB_SIZE = (120, 90)  # thumbnail size (w,h)
LINK_COLS = 3
LINK_PANEL_MAX_HEIGHT = 92

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SUPPORT_DIR = os.path.join(BASE_DIR, "supporting_documents")
VENDOR_LIST_PATH = os.path.join(SUPPORT_DIR, "vendor_list.txt")



# --------- Small helpers ---------
class ClickableLabel(QLabel):
    def __init__(self, text="", on_click=None, parent=None):
        super().__init__(text, parent)
        self.on_click = on_click
        self.setCursor(Qt.PointingHandCursor)

    def mousePressEvent(self, event):
        if callable(self.on_click):
            self.on_click()
        super().mousePressEvent(event)


class LinkLabel(ClickableLabel):
    """Blue link-looking label, underline on hover."""
    def __init__(self, text="", on_click=None, parent=None):
        super().__init__(text, on_click, parent)
        self._base_css = "color:#0a61d0; text-decoration:none; font-weight:500;"
        self._hover_css = "color:#0a61d0; text-decoration:underline; font-weight:500;"
        self.setStyleSheet(self._base_css)

    def enterEvent(self, e):
        self.setStyleSheet(self._hover_css)
        super().enterEvent(e)

    def leaveEvent(self, e):
        self.setStyleSheet(self._base_css)
        super().leaveEvent(e)


# --------- Main Tab ---------
class VendorQuoteTab(QWidget):
    def __init__(self, dirty_tracker=None, parent=None):
        super().__init__(parent)
        self.dirty_tracker = dirty_tracker
        self.vendor_rows = []

        self.vendor_list_path = VENDOR_LIST_PATH
        self._vendor_names = []
        self._vendor_labels = []
        self._ensure_vendor_file_exists()

        self.init_ui()

    # ---------- UI ----------
    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(0)

        # ---- Header row: Add Blank + hint text ----
        header = QWidget()
        header.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(10)

        btn_add_blank = QPushButton("Add Blank")
        btn_add_blank.setFixedHeight(26)
        btn_add_blank.setStyleSheet("font-weight: bold;")
        btn_add_blank.clicked.connect(self.add_vendor_row)

        hint = QLabel("Click vendor name to add frame")
        hint.setStyleSheet("color:#666; font-style: italic;")
        hint.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)

        header_layout.addWidget(btn_add_blank)
        header_layout.addWidget(hint)
        header_layout.addStretch()
        layout.addWidget(header)
        layout.addSpacing(2)

        # ---- Vendor "links" panel (scrollable) ----
        quick_panel = QWidget()
        quick_panel.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        qp_layout = QVBoxLayout(quick_panel)
        qp_layout.setContentsMargins(0, 0, 0, 0)
        qp_layout.setSpacing(0)

        self.link_container = QWidget()
        self.link_container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        pal = self.link_container.palette()
        pal.setColor(QPalette.ColorRole.Window, QColor("white"))
        self.link_container.setAutoFillBackground(True)
        self.link_container.setPalette(pal)

        self.link_grid = QGridLayout(self.link_container)
        self.link_grid.setContentsMargins(0, 0, 0, 0)
        self.link_grid.setHorizontalSpacing(18)
        self.link_grid.setVerticalSpacing(4)

        link_scroll = QScrollArea()
        link_scroll.setWidget(self.link_container)
        link_scroll.setWidgetResizable(True)
        link_scroll.setFrameShape(QFrame.NoFrame)
        link_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        link_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        link_scroll.setFixedHeight(LINK_PANEL_MAX_HEIGHT)
        self.link_scroll = link_scroll

        qp_layout.addWidget(link_scroll)
        layout.addWidget(quick_panel)

        # Build vendor link labels
        self._load_vendor_list()
        self._rebuild_vendor_links()

        # ---- Rows scroll area (gets remainder) ----
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setStyleSheet("QScrollArea { border:0; background: transparent; }")
        self.scroll.viewport().setStyleSheet("background: transparent;")

        self.row_container = QWidget()
        self.row_container.setStyleSheet("background: white;")
        self.row_container_layout = QVBoxLayout(self.row_container)
        self.row_container_layout.setContentsMargins(0, 0, 0, 0)
        self.row_container_layout.setSpacing(6)
        self.row_container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)

        self.scroll.setWidget(self.row_container)
        layout.addWidget(self.scroll, 1)

        # NOTE: No default blank row on launch.

    # ---------- Vendor list helpers ----------
    def _ensure_vendor_file_exists(self):
        if os.path.exists(self.vendor_list_path):
            return
        folder = os.path.dirname(self.vendor_list_path)
        if folder and not os.path.exists(folder):
            os.makedirs(folder, exist_ok=True)
        with open(self.vendor_list_path, "w", encoding="utf-8") as f:
            f.write("# One vendor per line. Lines beginning with # are ignored.\n"
                    "Acme Die\nBeta Converting\nCutRight Tooling\n")

    def _load_vendor_list(self):
        names = []
        try:
            with open(self.vendor_list_path, "r", encoding="utf-8-sig") as f:
                for line in f:
                    s = line.strip()
                    if not s or s.startswith("#"):
                        continue
                    names.append(s)
        except Exception:
            names = []

        seen = set()
        ordered = []
        for n in sorted(names, key=str.casefold):
            key = n.casefold()
            if key in seen:
                continue
            seen.add(key)
            ordered.append(n)
        self._vendor_names = ordered

    def _rebuild_vendor_links(self):
        # clear old
        while self.link_grid.count():
            item = self.link_grid.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self._vendor_labels.clear()

        r = c = 0
        for name in self._vendor_names:
            lbl = LinkLabel(name, on_click=lambda n=name: self.add_vendor_by_name(n))
            self._vendor_labels.append(lbl)
            self.link_grid.addWidget(lbl, r, c, Qt.AlignLeft)
            c += 1
            if c >= LINK_COLS:
                c = 0
                r += 1
        self.link_container.adjustSize()

        # auto-size panel height up to the cap
        count = len(self._vendor_labels)
        rows = (count + LINK_COLS - 1) // LINK_COLS
        needed_h = 0
        if rows > 0:
            row_h = self._vendor_labels[0].sizeHint().height()
            vspace = self.link_grid.verticalSpacing() or 0
            needed_h = rows * row_h + max(0, rows - 1) * vspace + 2
        self.link_scroll.setFixedHeight(min(LINK_PANEL_MAX_HEIGHT, max(0, needed_h)))

    # ---------- Add by name (prevents duplicates) ----------
    def add_vendor_by_name(self, name: str):
        existing_map = {}
        for row in self.vendor_rows:
            n = row.name_entry.text().strip()
            if n:
                existing_map[n.casefold()] = row

        key = name.casefold()
        if key in existing_map:
            # scroll to and flash existing
            row = existing_map[key]
            try:
                self.scroll.ensureWidgetVisible(row)
            except Exception:
                pass
            row.flash_card()
            return

        self.add_vendor_row(data=(name, "", []))

    # ---------- Rows API ----------
    def add_vendor_row(self, data=None):
        row = VendorQuoteRow(self, self.dirty_tracker, data)
        self.vendor_rows.append(row)
        self.row_container_layout.addWidget(row)

        sb = self.scroll.verticalScrollBar()
        sb.setValue(sb.maximum())

        if self.dirty_tracker and data is None:
            self.dirty_tracker.mark_dirty()

    def remove_vendor_row(self, row_widget):
        self.vendor_rows.remove(row_widget)
        row_widget.setParent(None)
        row_widget.deleteLater()
        if self.dirty_tracker:
            self.dirty_tracker.mark_dirty()

    def get_vendor_quote_data(self):
        return [row.get_row_data() for row in self.vendor_rows]

    def load_vendor_quote_data(self, data):
        self.clear_vendor_quote_tab(skip_add=True)
        for row_data in data:
            self.add_vendor_row(row_data)

    def clear_vendor_quote_tab(self, skip_add=False):
        for row in self.vendor_rows:
            row.setParent(None)
            row.deleteLater()
        self.vendor_rows.clear()
        # no default row added here


# --------- Row/Card ---------
class VendorQuoteRow(QWidget):
    def __init__(self, tab: QWidget, dirty_tracker=None, data=None):
        super().__init__()
        self.tab = tab
        self.dirty_tracker = dirty_tracker
        self._loading = bool(data)
        self.screenshots = []

        # card frame
        outer_frame = QFrame(self)
        outer_frame.setObjectName("VendorQuoteFrame")
        outer_frame.setFrameShape(QFrame.StyledPanel)
        outer_frame.setFrameShadow(QFrame.Raised)
        outer_frame.setStyleSheet("""
            QFrame#VendorQuoteFrame {
                border: 1.2px solid #c5c5c5;
                border-radius: 10px;
                background: #fafbfc;
            }
        """)
        self.outer_frame = outer_frame

        outer_layout = QVBoxLayout(outer_frame)
        outer_layout.setContentsMargins(12, 8, 12, 8)
        outer_layout.setSpacing(6)

        self.setLayout(QVBoxLayout())
        self.layout().setContentsMargins(0, 0, 0, 0)
        self.layout().addWidget(outer_frame)
        layout = outer_layout

        # top row
        row_top = QWidget()
        row_top_layout = QHBoxLayout(row_top)
        row_top_layout.setContentsMargins(0, 0, 0, 0)
        row_top_layout.setSpacing(6)

        self.btn_remove = QPushButton("Remove")
        self.btn_remove.setStyleSheet("""
            QPushButton {
                background: #ffeaea;
                color: #a10000;
                border: 1.0px solid #e0a4a4;
                border-radius: 6px;
                font-weight: 600;
                padding: 3px 10px;
            }
            QPushButton:hover { background: #fff0f0; }
        """)
        self.btn_remove.clicked.connect(lambda: self.tab.remove_vendor_row(self))
        row_top_layout.addWidget(self.btn_remove)

        self.name_entry = QLineEdit()
        self.name_entry.setPlaceholderText("Vendor Name")
        self.name_entry.setFixedWidth(180)
        self.name_entry.setFixedHeight(22)
        self.name_entry.textChanged.connect(self._on_user_change)
        row_top_layout.addWidget(self.name_entry)

        btn_paste = QPushButton("Paste Screenshot")
        btn_paste.setFixedHeight(22)
        btn_paste.clicked.connect(self.paste_screenshot)
        row_top_layout.addWidget(btn_paste)

        row_top_layout.addStretch()
        layout.addWidget(row_top)

        # notes box
        self.quote_text = QTextEdit()
        self.quote_text.setAcceptRichText(False)
        self.quote_text.setPlaceholderText("Enter quote/comments here (text only)")
        self.quote_text.setStyleSheet(
            "QTextEdit { border: 1px solid #ccc; border-radius: 3px; padding: 2px; }"
        )
        self.quote_text.document().setDocumentMargin(2)
        fm = self.quote_text.fontMetrics()
        self.min_textedit_height = fm.height() * 2 + 4
        self.quote_text.setMinimumHeight(self.min_textedit_height)
        self.quote_text.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.quote_text.textChanged.connect(self._on_textedit_changed)
        self.quote_text.installEventFilter(self)
        layout.addWidget(self.quote_text)

        # screenshots strip
        self.screenshot_container = QWidget()
        self.screenshot_layout = QHBoxLayout(self.screenshot_container)
        self.screenshot_layout.setAlignment(Qt.AlignLeft)
        self.screenshot_layout.setSpacing(8)
        self.screenshot_layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.screenshot_container)

        # populate if loading existing data
        if data:
            if data[0] and data[0] != self.name_entry.text():
                self.name_entry.setText(data[0])
            if len(data) > 1 and data[1] and data[1] != self.quote_text.toPlainText():
                self.quote_text.setPlainText(data[1])
                QTimer.singleShot(0, self.autosize_textedit)
            if len(data) > 2 and data[2]:
                self.load_screenshots(data[2])
        self._loading = False

    # quick flash when duplicate selected
    def flash_card(self):
        orig = self.outer_frame.styleSheet()
        self.outer_frame.setStyleSheet("""
            QFrame#VendorQuoteFrame {
                border: 2px solid #8ab4f8;
                border-radius: 10px;
                background: #fffbd1;
            }
        """)
        QTimer.singleShot(650, lambda: self.outer_frame.setStyleSheet(orig))

    # autosize / events
    def eventFilter(self, obj, ev):
        if obj is self.quote_text and ev.type() == QEvent.Resize:
            self.autosize_textedit()
        return super().eventFilter(obj, ev)

    def _on_user_change(self):
        if not self._loading and self.dirty_tracker:
            self.dirty_tracker.mark_dirty()

    def _on_textedit_changed(self):
        self.autosize_textedit()
        self._on_user_change()

    def autosize_textedit(self):
        doc = self.quote_text.document()
        doc.setTextWidth(self.quote_text.viewport().width())
        doc_h = doc.documentLayout().documentSize().height()
        m = self.quote_text.contentsMargins()
        frame = self.quote_text.frameWidth()
        padding = 6
        new_h = max(self.min_textedit_height,
                    int(doc_h + m.top() + m.bottom() + frame * 2 + padding))
        self.quote_text.setFixedHeight(new_h)

    # screenshots
    def paste_screenshot(self):
        clipboard = QGuiApplication.clipboard()
        if clipboard.mimeData().hasImage():
            img = clipboard.image()
            pixmap = QPixmap.fromImage(img)
            b64 = self.encode_qimage_base64(img)
            self.add_screenshot_entry(b64, pixmap=pixmap, lazy=False)
            self._on_user_change()
        else:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "No Image", "Clipboard does not contain an image.")

    def encode_qimage_base64(self, img):
        if hasattr(QImage, "Format_RGBA8888"):
            img = img.convertToFormat(QImage.Format_RGBA8888)
        else:
            img = img.convertToFormat(QImage.Format_ARGB32)
        buf = QBuffer()
        buf.open(QBuffer.WriteOnly)
        img.save(buf, "PNG")
        return base64.b64encode(buf.data()).decode("utf-8")

    def add_screenshot_entry(self, b64, pixmap=None, lazy=False):
        thumb_widget = QWidget()
        vbox = QVBoxLayout(thumb_widget)
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.setSpacing(3)

        if not lazy and pixmap is not None:
            thumb = pixmap.scaled(
                THUMB_SIZE[0], THUMB_SIZE[1],
                Qt.KeepAspectRatio, Qt.SmoothTransformation
            )
            label = ClickableLabel()
            label.setPixmap(thumb)
            label.setAlignment(Qt.AlignCenter)
            label.setFixedSize(THUMB_SIZE[0], THUMB_SIZE[1])
            label.setCursor(Qt.PointingHandCursor)
            label.setToolTip("Click to view full size")
            label.on_click = lambda pixmap=pixmap: self.show_fullsize_screenshot(pixmap)
            vbox.addWidget(label, alignment=Qt.AlignCenter)

        btn_row = QWidget()
        btn_layout = QHBoxLayout(btn_row)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(4)

        btn_view = QPushButton("View")
        btn_view.setFixedWidth(54)
        if lazy:
            btn_view.clicked.connect(lambda: self.load_and_show_screenshot(b64, thumb_widget))
        else:
            btn_view.clicked.connect(lambda: self.show_fullsize_screenshot(pixmap))
        btn_layout.addWidget(btn_view)

        btn_delete = QPushButton("Delete")
        btn_delete.setFixedWidth(54)
        btn_delete.clicked.connect(lambda: self.remove_screenshot(thumb_widget, b64))
        btn_layout.addWidget(btn_delete)

        vbox.addWidget(btn_row, alignment=Qt.AlignCenter)
        self.screenshot_layout.addWidget(thumb_widget)
        self.screenshots.append((thumb_widget, b64))

    def load_and_show_screenshot(self, b64, _container_widget):
        img_bytes = base64.b64decode(b64)
        pixmap = QPixmap()
        pixmap.loadFromData(img_bytes)
        self.show_fullsize_screenshot(pixmap)

    def remove_screenshot(self, thumb_widget, b64):
        for widget, encoded in list(self.screenshots):
            if widget == thumb_widget and encoded == b64:
                self.screenshot_layout.removeWidget(widget)
                widget.deleteLater()
                self.screenshots.remove((widget, b64))
                break
        self._on_user_change()

    def load_screenshots(self, b64_list):
        self.screenshot_container.setUpdatesEnabled(False)
        try:
            for b64 in b64_list:
                self.add_screenshot_entry(b64, lazy=True)
        finally:
            self.screenshot_container.setUpdatesEnabled(True)

    def show_fullsize_screenshot(self, pixmap):
        class ClickToCloseLabel(QLabel):
            def __init__(self, dialog):
                super().__init__()
                self.dialog = dialog
            def mousePressEvent(self, event):
                self.dialog.accept()

        dlg = QDialog(self)
        dlg.setWindowTitle("Screenshot (Full Size)")
        dlg.setModal(True)

        layout = QVBoxLayout(dlg)
        scroll = QScrollArea(dlg)
        label = ClickToCloseLabel(dlg)
        label.setPixmap(pixmap)
        label.setAlignment(Qt.AlignCenter)
        label.setCursor(Qt.PointingHandCursor)
        label.setToolTip("Click image to close")
        scroll.setWidget(label)
        scroll.setWidgetResizable(True)
        layout.addWidget(scroll)

        hint = QLabel("<span style='font-size:15px; font-weight:bold; text-decoration:underline;'>Click image to close</span>")
        hint.setAlignment(Qt.AlignCenter)
        layout.addWidget(hint)

        screen = QGuiApplication.primaryScreen().availableGeometry()
        dlg.resize(min(pixmap.width() + 40, int(screen.width() * 0.9)),
                   min(pixmap.height() + 60, int(screen.height() * 0.9)))
        dlg.exec()

    def get_row_data(self):
        name = self.name_entry.text().strip()
        text = self.quote_text.toPlainText().strip()
        imgs = [b64 for (_, b64) in self.screenshots]
        return (name, text, imgs)

    def showEvent(self, event):
        super().showEvent(event)
        self.autosize_textedit()


# ---- Factory helpers used by launch.py ----
_tab_instance = None

def create_vendor_quote_tab(tab_widget: QWidget, dirty_tracker=None):
    global _tab_instance
    _tab_instance = VendorQuoteTab(dirty_tracker)
    layout = QVBoxLayout(tab_widget)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.addWidget(_tab_instance)

def get_vendor_quote_data():
    return _tab_instance.get_vendor_quote_data() if _tab_instance else []

def load_vendor_quote_data(data):
    if _tab_instance:
        _tab_instance.load_vendor_quote_data(data)

def clear_vendor_quote_tab(skip_add=False):
    if _tab_instance:
        _tab_instance.clear_vendor_quote_tab(skip_add=skip_add)
