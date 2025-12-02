import os
import json
import re
from datetime import datetime

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QLineEdit, QLabel, QSizePolicy, QMessageBox,
    QAbstractItemView, QHeaderView
)
from PySide6.QtCore import Qt, QTimer

# --------- paths ---------
# Match launch.py
CHECKLISTS_DIR = r"P:\ENGINEERING\Design Checklist\json_files"

HEADERS = ["ID #", "Customer", "Filename", "Date", "Last Edited By"]

# Per-user settings file (width persistence)
APPDATA_DIR = os.environ.get("APPDATA") or os.path.expanduser("~")
SETTINGS_DIR = os.path.join(APPDATA_DIR, "EngineeringChecklist")
WIDTHS_PATH = os.path.join(SETTINGS_DIR, "saved_tab_widths.json")

# Keep a global instance so launch.py helpers can interact with the tab
_tab_instance = None


# ---------- Sorting-friendly table item ----------
class SortItem(QTableWidgetItem):
    """
    QTableWidgetItem that sorts by an explicit sort_key (numeric/date-safe),
    while displaying a human-friendly text.
    """
    def __init__(self, text: str, sort_key=None):
        super().__init__(text)
        # Always store a comparable sort_key (no type mixing)
        if isinstance(sort_key, tuple):
            self._sort_key = sort_key
        elif sort_key is None:
            self._sort_key = (1, (text or "").lower())
        elif isinstance(sort_key, (int, float)):
            self._sort_key = (0, sort_key)
        else:
            self._sort_key = (1, str(sort_key).lower())

    def __lt__(self, other):
        if isinstance(other, SortItem):
            return self._sort_key < other._sort_key
        return super().__lt__(other)


def _extract_id_from_top_fields(top_fields):
    """
    Return the numeric ID (4+ digits) if present in any of the top fields.
    If not found, return "" (do NOT guess from any fixed index).
    Works for both legacy [Customer, Opp, ID, Sales] and new [ID, Customer, Opp, Sales].
    """
    if not isinstance(top_fields, (list, tuple)):
        return ""
    for val in top_fields:
        if isinstance(val, str):
            s = val.strip()
            if re.fullmatch(r"\d{4,}", s):
                return s
    return ""


class SavedChecklistsTab(QWidget):
    def __init__(self, load_checklist_callback=None, parent=None):
        super().__init__(parent)
        self.load_checklist_callback = load_checklist_callback

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(6)

        # Top bar: search + refresh + open
        top = QWidget()
        top_l = QHBoxLayout(top)
        top_l.setContentsMargins(0, 0, 0, 0)
        top_l.setSpacing(6)

        self.search = QLineEdit()
        self.search.setPlaceholderText("Search filename…")

        self.btn_refresh = QPushButton("Refresh")
        self.btn_open = QPushButton("Open")

        top_l.addWidget(QLabel("Filter:"))
        top_l.addWidget(self.search)
        top_l.addStretch()
        top_l.addWidget(self.btn_refresh)
        top_l.addWidget(self.btn_open)
        layout.addWidget(top)

        # Table
        self.table = QTableWidget(0, len(HEADERS))
        self.table.setHorizontalHeaderLabels(HEADERS)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table.doubleClicked.connect(self.open_selected)

        # Selection behavior (entire row); NO hover highlight, stable selection color
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setStyleSheet(
            "QTableView::item:hover { background: transparent; }"
            "QTableView::item:selected { background: #e6f2ff; color: black; }"
        )

        # Column sizing: all interactive (user-resizable). No default stretch.
        hdr = self.table.horizontalHeader()
        hdr.setStretchLastSection(False)
        for col in range(len(HEADERS)):
            hdr.setSectionResizeMode(col, QHeaderView.ResizeMode.Interactive)

        # Sorting enabled (we'll default to Date ↓ after populate)
        self.table.setSortingEnabled(True)

        layout.addWidget(self.table, 1)

        # --------- cache + debounce ---------
        self._entries = []  # cached (customer, filename, mtime, user, id_val)

        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(lambda: self.update_table(rescan=False))

        self.search.textChanged.connect(self._on_search_changed)
        self.btn_refresh.clicked.connect(lambda: self.update_table(rescan=True))
        self.btn_open.clicked.connect(self.open_selected)

        # Width persistence flags
        self._has_saved_widths = False
        self._widths_applied = False
        self._load_saved_widths()

        # Initial populate (scan once)
        self.update_table(rescan=True)

    # ---------- Per-user column widths ----------
    def _load_saved_widths(self):
        try:
            if os.path.exists(WIDTHS_PATH):
                with open(WIDTHS_PATH, "r", encoding="utf-8") as f:
                    obj = json.load(f)
                widths = obj.get("saved_tab_column_widths")
                if isinstance(widths, list) and len(widths) == len(HEADERS):
                    self._saved_widths = widths
                    self._has_saved_widths = True
                else:
                    self._saved_widths = []
                    self._has_saved_widths = False
            else:
                self._saved_widths = []
                self._has_saved_widths = False
        except Exception:
            self._saved_widths = []
            self._has_saved_widths = False

    def apply_column_widths(self):
        """Apply saved widths once per session (after first populate)."""
        if self._widths_applied:
            return
        if self._has_saved_widths:
            for i, w in enumerate(self._saved_widths):
                try:
                    if isinstance(w, int) and w > 10:
                        self.table.setColumnWidth(i, w)
                except Exception:
                    pass
        else:
            # No saved widths: give helpful starting widths just once
            self.table.resizeColumnToContents(0)  # ID
            self.table.resizeColumnToContents(1)  # Customer
            self.table.resizeColumnToContents(3)  # Date
            self.table.resizeColumnToContents(4)  # Last Edited By
        self._widths_applied = True

    def save_column_widths(self):
        r"""Persist current widths to %APPDATA%\EngineeringChecklist\saved_tab_widths.json"""
        try:
            widths = [self.table.columnWidth(i) for i in range(len(HEADERS))]
            os.makedirs(SETTINGS_DIR, exist_ok=True)
            obj = {"saved_tab_column_widths": widths}
            with open(WIDTHS_PATH, "w", encoding="utf-8") as f:
                json.dump(obj, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[SavedTab] Failed to save column widths: {e}")

    # ---------- Fast scanner ----------
    def list_files(self):
        """
        Return list of tuples:
          (customer, filename, mtime, user, id_val)
        Scans the directory with os.scandir() and parses each JSON once.
        """
        entries = []

        # If the directory doesn't exist (e.g., P:\ is gone), return empty list
        if not os.path.isdir(CHECKLISTS_DIR):
            return []

        try:
            with os.scandir(CHECKLISTS_DIR) as it:
                for de in it:
                    if not de.is_file():
                        continue
                    name = de.name
                    if not name.lower().endswith(".json"):
                        continue

                    # Cheap mtime
                    try:
                        mtime = de.stat().st_mtime
                    except Exception:
                        mtime = 0.0

                    # Parse once
                    customer, user, id_val = "", "", ""
                    try:
                        with open(de.path, "r", encoding="utf-8") as f:
                            data = json.load(f)

                        user = data.get("last_user", "") or ""
                        top = (data.get("checklist") or {}).get("top_fields", []) or []

                        # Customer: first non-empty non-numeric field
                        if isinstance(top, list):
                            for v in top:
                                if isinstance(v, str) and v.strip() and not v.strip().isdigit():
                                    customer = v.strip()
                                    break

                        # ID detection (numeric only)
                        id_val = _extract_id_from_top_fields(top)

                    except Exception:
                        pass  # keep defaults

                    entries.append((customer, name, mtime, user, id_val))

            # newest first
            entries.sort(key=lambda x: x[2], reverse=True)
            return entries

        except Exception:
            # No console error; just return an empty list
            return []


    # ---------- UI wiring ----------
    def _on_search_changed(self, _text: str):
        # debounce: wait 150ms after the last keystroke
        self._search_timer.start(150)

    def update_table(self, rescan=True):
        """
        Populate the table. If rescan=True, rescan the folder and rebuild the cache.
        Otherwise, filter the cached entries in memory (fast).
        """
        # Rebuild cache only when asked (first load / Refresh)
        if rescan or not self._entries:
            self._entries = self.list_files()

        q = (self.search.text() or "").strip().lower()

        def _match(cust, fname, id_val):
            name_no_ext = fname[:-5] if fname.lower().endswith(".json") else fname
            return (
                (q in (id_val or "").lower()) or
                (q in (cust   or "").lower()) or
                (q in (name_no_ext or "").lower())
            )

        files = self._entries
        filtered = [
            (cust, fname, mtime, user, id_val)
            for (cust, fname, mtime, user, id_val) in files
            if (not q) or _match(cust, fname, id_val)
        ]

        # Temporarily turn sorting off while we repopulate to avoid flicker
        was_sorting = self.table.isSortingEnabled()
        if was_sorting:
            self.table.setSortingEnabled(False)

        self.table.setRowCount(len(filtered))

        for row, (cust, fname, mtime, user, id_val) in enumerate(filtered):
            display = fname[:-5] if fname.lower().endswith(".json") else fname
            date_str = datetime.fromtimestamp(mtime).strftime("%m/%d/%Y")

            # Uniform sort keys
            try:
                id_sort = int(id_val) if id_val and id_val.isdigit() else -1
            except Exception:
                id_sort = -1

            items = [
                SortItem(id_val, sort_key=id_sort),                  # ID (numeric sort)
                SortItem(cust, sort_key=(cust or "").lower()),       # Customer (alpha)
                SortItem(display, sort_key=(display or "").lower()), # Filename (alpha)
                SortItem(date_str, sort_key=mtime),                  # Date (numeric mtime)
                SortItem(user, sort_key=(user or "").lower()),       # Last Edited By (alpha)
            ]
            for col, it in enumerate(items):
                it.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(row, col, it)

        # Apply saved widths (or first-time helpful widths) once per session
        self.apply_column_widths()

        # Restore sorting and set default to Date ↓ (newest first)
        self.table.setSortingEnabled(True)
        self.table.sortItems(3, Qt.SortOrder.DescendingOrder)

    def open_selected(self):
        sel = self.table.currentRow()
        if sel < 0:
            QMessageBox.information(self, "Open", "Please select a checklist to open.")
            return
        display_name = self.table.item(sel, 2).text()  # "Filename" column (no .json)
        fname = f"{display_name}.json"
        path = os.path.join(CHECKLISTS_DIR, fname)
        if not os.path.exists(path):
            QMessageBox.warning(self, "Missing File", f"File not found:\n{path}")
            return
        if self.load_checklist_callback:
            self.load_checklist_callback(path)


# ========== Launch.py compatibility helpers ==========

def create_saved_checklists_tab(tab_widget: QWidget, load_checklist_callback=None, notebook=None):
    """Factory to match launch.py's expectation."""
    global _tab_instance
    _tab_instance = SavedChecklistsTab(load_checklist_callback=load_checklist_callback)
    lay = QVBoxLayout(tab_widget)
    lay.setContentsMargins(0, 0, 0, 0)
    lay.addWidget(_tab_instance)

def clear_saved_checklists_tab():
    """Called by launch.new_checklist_action()."""
    if _tab_instance:
        _tab_instance.update_table(rescan=True)

def save_column_widths_global():
    """launch.on_close() calls this to persist widths."""
    if _tab_instance:
        _tab_instance.save_column_widths()
