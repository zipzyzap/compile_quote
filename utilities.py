import os
import re
import time
import getpass
from PySide6.QtWidgets import QMessageBox

def require_file(path, parent=None, description="file"):
    """
    Ensure a file/folder exists before using it.
    Shows a QMessageBox if missing and returns False; otherwise True.
    """
    if not path or not os.path.exists(path):
        QMessageBox.warning(
            parent,
            "File Not Found",
            f"The required {description} was not found:\n\n{path}"
        )
        return False
    return True

# ----------- Constants -----------
STRATASYS_PRICES = {
    "515 ABS": 0.3465,
    "531 ABS": 0.3465,
    "706 Sup": 0.12075,
    "810 VC": 0.3465,
    "820 VUC": 0.3465,
    "825 VUW": 0.3465,
    "865 VUB": 0.3465,
    "985 AB": 0.3465
}
FORMLABS_PRICE = 0.349

STRATASYS_ORDER = [
    "985 AB",
    "515 ABS",
    "531 ABS",
    "865 VUB",
    "825 VUW",
    "820 VUC",
    "810 VC",
    "706 Sup"
]
FORMLABS_HEADERS = ["RS-F2", "Time (hrs)", "Material $", "3D Cost"]

# ---- Cached index for drawings folder ----
_DRAWINGS_CACHE = {"path": None, "mtime": None, "index": None}

def _normalize_root(name: str) -> str:
    # Match how the rest of the app normalizes: upper, spaces/dashes -> underscores
    return (name or "").strip().upper().replace(" ", "_").replace("-", "_")

def _build_drawings_index(drawings_folder="P:/PDF Drawings"):
    """
    Build an index: { ROOT -> [(rev_str, filename), ...] }
    Where filename is the actual filename in the folder.
    """
    import os, re
    if not require_file(drawings_folder, description="PDF drawings folder"):
        return {}
    try:
        files = os.listdir(drawings_folder)
    except Exception:
        # If something goes wrong listing the folder, treat as empty
        return {}
    # Example filename patterns we’ve seen:  MT29941_Rev0.pdf, AB12345_Rev1.2 XYZ.PDF
    pat = re.compile(r'^(?P<root>.+?)_Rev(?P<rev>[0-9]+(?:\.[0-9]+)?)', re.IGNORECASE)
    index = {}
    for name in files:
        m = pat.match(name)
        if not m:
            continue
        root = _normalize_root(m.group("root"))
        rev  = m.group("rev")
        index.setdefault(root, []).append((rev, name))
    return index

def _get_drawings_index(drawings_folder="P:/PDF Drawings"):
    """
    Return cached index; rebuild if folder path/mtime changes or cache empty.
    """
    import os
    try:
        mtime = os.path.getmtime(drawings_folder)
    except Exception:
        # If the share isn’t available, leave cache as-is and return empty
        return {}

    if (_DRAWINGS_CACHE["path"] != drawings_folder or
        _DRAWINGS_CACHE["mtime"] != mtime or
        _DRAWINGS_CACHE["index"] is None):
        index = _build_drawings_index(drawings_folder)
        _DRAWINGS_CACHE.update({"path": drawings_folder, "mtime": mtime, "index": index})
    return _DRAWINGS_CACHE["index"]


# ----------- File/Revision/Format Utilities -----------

def extract_numeric_part(text):
    """Extract the first integer found in text (or inf if none)."""
    match = re.search(r'(\d+)', text)
    return int(match.group(1)) if match else float('inf')

def extract_revision(filename):
    """Extracts revision string from filename."""
    match = re.search(r'_rev([0-9]+(?:\.[0-9]+)?)', filename, re.IGNORECASE)
    return match.group(1) if match else None

def format_rev(rev):
    """Formats a revision string for filenames."""
    rev = rev.rstrip(".")
    return f"Rev{rev}"

def format_drawing_with_rev(tup):
    """Formats drawing and revision for filenames."""
    drawing, rev = tup
    drawing = drawing.upper().replace(" ", "_").replace("-", "_")
    return f"{drawing}_{format_rev(rev)}"

def rev_key(rev):
    """Converts a revision string like '0', '0.1' into a sortable tuple."""
    if not rev:
        return (0,)
    return tuple(int(p) if p.isdigit() else 0 for p in rev.split("."))

def find_latest_revision_files(drawing, drawings_folder="P:/PDF Drawings"):
    """
    Return dict {rev_str: filename} for all files that match the given drawing root.
    Uses a cached index to avoid re-listing the directory.
    """
    root = _normalize_root(drawing)
    idx = _get_drawings_index(drawings_folder)
    if not idx:
        return {}
    pairs = idx.get(root, [])
    # Keep the same return shape you already depend on: {rev: filename}
    return {rev: filename for (rev, filename) in pairs}

def find_latest_pdf_with_rev(part_number, drawings_folder="P:/PDF Drawings"):
    """Returns (path, display_name) of the latest PDF for the part_number."""
    files = [
        f for f in os.listdir(drawings_folder)
        if f.lower().endswith('.pdf') and part_number.upper() in f.upper()
    ]
    if not files:
        return None, part_number
    def pdf_rev_key(f):
        m = re.search(r'_Rev([\d\.]+)', f, re.IGNORECASE)
        rev = m.group(1) if m else "0"
        return rev_key(rev)
    files.sort(key=pdf_rev_key)
    newest = files[-1]
    full_path = os.path.join(drawings_folder, newest)
    display_name = os.path.splitext(newest)[0]
    return full_path, display_name

def get_all_part_numbers_and_revs(quote_info_rows):
    """
    Returns a list of formatted part numbers with their latest revision,
    using all rows in the Quote Info tab.
    """
    found = []
    for row in quote_info_rows:
        fields = row.get("fields", [])
        if not fields or not fields[0]:
            continue
        part_number = fields[0].strip()
        if not part_number:
            continue
        matched = find_latest_revision_files(part_number)
        if not matched:
            continue
        latest_rev = max(matched.keys(), key=rev_key)
        formatted = format_drawing_with_rev((part_number, latest_rev))
        found.append(formatted)

    found = sorted(set(found), key=lambda x: [int(s) if s.isdigit() else s for s in x.replace("_Rev", " ").split()])
    return found

# ----------- 3D Quote Calculation Helpers -----------

def calculate_stratasys_cost(s_vars):
    total = 0
    for m, price in STRATASYS_PRICES.items():
        try:
            qty = float(s_vars.get(m, "0"))
            total += qty * price
        except Exception:
            continue
    total *= 1.3
    return total

def calculate_stratasys_3d_cost(material_total, time_hrs):
    try:
        t = float(time_hrs)
    except Exception:
        t = 0
    return material_total + t * 20

def calculate_formlabs_cost(f_vars):
    try:
        qty = float(f_vars.get("RS-F2", "0"))
    except Exception:
        qty = 0
    mat = qty * FORMLABS_PRICE * 1.3
    return mat

def calculate_formlabs_3d_cost(mat, t_f):
    try:
        t_f = float(t_f)
    except Exception:
        t_f = 0
    return mat + t_f * 20

class DirtyTracker:
    def __init__(self):
        self._dirty = False
        self._callback = None

    def set_callback(self, callback):
        self._callback = callback

    def is_dirty(self):
        return self._dirty

    def mark_dirty(self):
        if not self._dirty:
            self._dirty = True
            if self._callback:
                self._callback()

    def mark_clean(self):
        if self._dirty:
            self._dirty = False
            if self._callback:
                self._callback()

# ----------- Checklist Lock File Helpers -----------

def get_lock_path(json_path):
    return json_path + ".lock"

def lock_checklist(json_path):
    lock_path = get_lock_path(json_path)
    if os.path.exists(lock_path):
        try:
            with open(lock_path, "r") as f:
                locker = f.read().strip()
        except Exception:
            locker = None
        return False, locker  # Already locked
    else:
        # Lock it for the current user
        with open(lock_path, "w") as f:
            f.write(f"{getpass.getuser()} at {time.ctime()}")
        return True, None

def unlock_checklist(json_path):
    lock_path = get_lock_path(json_path)
    try:
        if os.path.exists(lock_path):
            os.remove(lock_path)
    except Exception as e:
        print(f"Error removing lock file: {e}")

def force_unlock_checklist(json_path):
    unlock_checklist(json_path)

def convert_inches_to_tonnage(inches: float) -> float:
    return inches * 0.0008  # Adjust multiplier as needed

def open_rotary_die_file():
    import subprocess
    rotary_path = r"P:\ENGINEERING\Design Checklist\Rotary Die List.xlsx"
    if os.path.exists(rotary_path):
        try:
            os.startfile(rotary_path)
        except Exception as e:
            print(f"Error opening file: {e}")
            