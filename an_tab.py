from PySide6.QtWidgets import QWidget, QVBoxLayout, QTextEdit

class AdditionalNotesTab(QWidget):
    def __init__(self, dirty_tracker=None, parent=None):
        super().__init__(parent)
        self.dirty_tracker = dirty_tracker
        self._loading = False

        self.notes_edit = QTextEdit()
        self.notes_edit.setPlaceholderText("Enter additional notes here...")
        self.notes_edit.textChanged.connect(self.on_notes_changed)
        layout = QVBoxLayout(self)
        layout.addWidget(self.notes_edit)

    def on_notes_changed(self):
        if self._loading or not self.dirty_tracker:
            return
        self.dirty_tracker.mark_dirty()

    def get_notes_text(self):
        return self.notes_edit.toPlainText()

    def set_notes_text(self, text):
        self._loading = True
        self.notes_edit.setPlainText(text or "")
        self._loading = False

    def set_read_only(self, read_only):
        self.notes_edit.setReadOnly(read_only)

_tab_instance = None

def create_additional_notes_tab(tab_widget: QWidget, dirty_tracker=None):
    global _tab_instance
    _tab_instance = AdditionalNotesTab(dirty_tracker)
    layout = QVBoxLayout(tab_widget)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.addWidget(_tab_instance)

def get_notes_text():
    return _tab_instance.get_notes_text() if _tab_instance else ""

def set_notes_text(text):
    if _tab_instance:
        _tab_instance.set_notes_text(text)

def set_read_only(read_only):
    if _tab_instance:
        _tab_instance.set_read_only(read_only)
