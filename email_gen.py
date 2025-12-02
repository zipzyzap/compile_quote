# email_gen.py — fast + reliable attachments (historic speed, your rules kept)
from __future__ import annotations
import re
from pathlib import Path
from typing import Dict, List, Tuple
import win32com.client as win32
from PySide6.QtWidgets import QMessageBox

from utilities import (
    find_latest_revision_files,
    extract_numeric_part,
    rev_key,
    format_drawing_with_rev,
)

ATTACH_ROOT = Path(r"P:\PDF Drawings")
ALLOWED_EXTS = [".pdf", ".step", ".dwg", ".dxf", ".igs"]
MAX_PAGE_IDX = 49  # supports _0.._49

def _normalize(s: str) -> str:
    return (s or "").strip()

def _ensure_part_with_rev(drawing: str) -> str:
    """
    Ensure part number is in the form {Part-Number}_RevX.
    MIS parts are always MIS-#### (hyphen, never underscore in part number).
    """
    drawing = drawing.strip()

    # Normalize MIS to always have a hyphen in the part number
    if drawing.upper().startswith("MIS_"):
        drawing = "MIS-" + drawing[4:]

    # If already has revision, just return
    if re.search(r"_Rev[\w.\-]+$", drawing, re.IGNORECASE):
        return drawing

    # Try to find latest revision from files
    matches = find_latest_revision_files(drawing) or {}
    if matches:
        latest = max(matches.keys(), key=rev_key)
        return f"{drawing}_Rev{latest}"

    return drawing


def _attachment_candidates(part_with_rev: str) -> List[Path]:
    """
    Build candidate file paths for the given part_with_rev.
    Always assumes MIS parts use hyphen in part number.
    """
    # Normalize MIS prefix just in case
    if part_with_rev.upper().startswith("MIS_"):
        part_with_rev = "MIS-" + part_with_rev[4:]

    cands: List[Path] = []
    for ext in ALLOWED_EXTS:
        for e in (ext, ext.upper()):
            cands.append(ATTACH_ROOT / f"{part_with_rev}{e}")  # exact
            for i in range(MAX_PAGE_IDX + 1):                   # paged
                cands.append(ATTACH_ROOT / f"{part_with_rev}_{i}{e}")

    # Deduplicate while preserving order
    seen = set()
    out = []
    for p in cands:
        key = p.as_posix().lower()
        if key not in seen:
            seen.add(key)
            out.append(p)
    return out


def _build_body(category: str, blocks: List[Tuple[str, str, str]]) -> str:
    if category == "CD":
        return (
            "Hi Paula/Jane,\n\n"
            "Please quote a flex die for the attached drawings. Tool layouts attached.\n\n"
            "Thank you"
        )
    header = (
        "Hello,\n\n"
        "Please quote price on the attached drawing and file. Remember to include weight in grams.\n\n"
        "Thank you\n\n"
    ) if category == "MT" else (
        "Hello\n\n"
        "Please see the attached MIS drawings for quote.\n"
        "Let me know if you have any questions.\n\n"
        "Thank you\n\n"
    )
    # MT/MIS: blank line between parts
    parts = [f"{p}\nMaterial: {m}\nQuantity: {q}" for (p,m,q) in blocks]
    return header + ("\n\n".join(parts) if parts else "")

def generate_emails(quote_data, parent=None):
    categorized: Dict[str, List[Tuple[str, str, str]]] = {"CD": [], "MT": [], "MIS": []}
    # Collect rows
    for row in quote_data:
        if not row.get("include", True): 
            continue
        f = row.get("fields", [])
        if not f or len(f) < 3: 
            continue
        drawing, material, quantity = f[0].strip(), f[1].strip(), f[2].strip()
        up = drawing.upper()
        if   up.startswith("CD"):  categorized["CD"].append((drawing, material, quantity))
        elif up.startswith("MT"):  categorized["MT"].append((drawing, material, quantity))
        elif up.startswith("MIS"): categorized["MIS"].append((drawing, material, quantity))

    try:
        outlook = win32.Dispatch("Outlook.Application")
    except Exception as e:
        QMessageBox.critical(parent, "Outlook Error", f"Failed to launch Outlook:\n{e}")
        return

    missing_any: List[str] = []
    for category, entries in categorized.items():
        if not entries:
            continue

        # Create & display first so Outlook inserts default signature
        mail = outlook.CreateItem(0)
        mail.Display()

        # Sort like historic behavior
        entries_sorted = sorted(entries, key=lambda x: extract_numeric_part(x[0]))

        subject_parts: List[str] = []
        body_blocks: List[Tuple[str,str,str]] = []

        for drawing, material, quantity in entries_sorted:
            part = _ensure_part_with_rev(drawing)          # e.g., CD27001_Rev1.3
            subject_parts.append(part)

            # Try to attach every candidate that exists
            attached = False
            for p in _attachment_candidates(part):
                try:
                    if p.exists():
                        mail.Attachments.Add(str(p.resolve()))
                        attached = True
                except Exception:
                    # keep going if a single add fails
                    pass
            if not attached:
                missing_any.append(part)

            if category in ("MT","MIS"):
                body_blocks.append((part, material, quantity))

        # Subject
        mail.Subject = f"RFQ: {' '.join(subject_parts)}" if subject_parts else "RFQ: No valid parts found"

        # Body
        body = _build_body(category, body_blocks)
        mail.HTMLBody = body.replace("\n", "<br>") + "<br><br>" + mail.HTMLBody

    if missing_any:
        QMessageBox.warning(
            parent,
            "Missing Attachments",
            "No files found for:\n" + "\n".join(sorted(set(missing_any)))
        )
