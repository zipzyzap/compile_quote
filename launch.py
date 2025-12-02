import os
import sys
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QTabWidget,
    QWidget,
    QVBoxLayout,
    QPushButton,
    QMessageBox,
    QFileDialog,
    QHBoxLayout,
    QScrollArea,
    QLabel,
    QDialog,
    QTextEdit,
    
)
from PySide6.QtGui import (
    QIcon,
    QKeySequence,
    QCursor,
    QDesktopServices
)
from PySide6.QtCore import Qt, QUrl

import user_settings
import cl_tab
import qi_tab
import ref_tab
import vq_tab
import html_export
import saved_tab
import an_tab
import cd_ref

from utilities import (
    find_latest_revision_files,
    format_drawing_with_rev,
    get_all_part_numbers_and_revs,
    lock_checklist,
    unlock_checklist,
    get_lock_path,
    DirtyTracker
)

APP_VERSION = "0.0"
CHECKLISTS_DIR = r"P:\ENGINEERING\Design Checklist\json_files"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_checklist_path = None

        # --- Dirty trackers ---
        self.dirty_trackers = {
            "Checklist": DirtyTracker(),
            "Quote Info": DirtyTracker(),
            "Vendor Quotes": DirtyTracker(),
            "Additional Notes": DirtyTracker(),
        }
        for tracker in self.dirty_trackers.values():
            tracker.set_callback(self.update_window_title)

        self.current_lockfile = None
        self.settings = user_settings.load_user_settings()
        self.setWindowTitle(f"Engineering Checklist {APP_VERSION}")

        # --- Window geometry ---
        self.resize(850, 800)
        self.setMinimumSize(850, 800)
        geometry = self.settings.get("window_geometry", None)
        if geometry:
            try:
                size_part, pos_part = geometry.split('+', 1)
                w, h = map(int, size_part.split('x'))
                self.resize(w, h)
            except Exception:
                pass

        # --- App icon ---
        icon_path = self.resource_path("appiconnew.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # --- Tabs setup ---
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        self.tab_names = [
            "Reference",
            "Die Cut Reference",
            "Checklist",
            "Quote Info",
            "Vendor Quotes",
            "Additional Notes",
            "Saved Checklists"
        ]
        self.tab_widgets = {}
        for name in self.tab_names:
            widget = QWidget()
            self.tabs.addTab(widget, name)
            self.tab_widgets[name] = widget

        # Reference tab
        ref_tab.create_reference_tab(self.tab_widgets["Reference"])

        # Rotary Reference (Gap Calculator) tab
        self.gap_calc_widget = cd_ref.GapCalculatorWidget()
        gap_layout = QVBoxLayout(self.tab_widgets["Die Cut Reference"])
        gap_layout.setContentsMargins(4, 4, 4, 4)
        gap_layout.setSpacing(2)
        gap_layout.addWidget(self.gap_calc_widget)

        # Checklist tab
        cl_tab.create_checklist_tab(
            self.tab_widgets["Checklist"],
            self.dirty_trackers["Checklist"]
        )
        self.checklist_tab = cl_tab._tab_instance

        # Quote Info tab
        qi_tab.create_quote_info_tab(
            self.tab_widgets["Quote Info"],
            self.dirty_trackers["Quote Info"]
        )

        # Vendor Quotes tab
        vq_tab.create_vendor_quote_tab(
            self.tab_widgets["Vendor Quotes"],
            self.dirty_trackers["Vendor Quotes"]
        )

        # Additional Notes tab
        an_tab.create_additional_notes_tab(
            self.tab_widgets["Additional Notes"],
            self.dirty_trackers["Additional Notes"]
        )

        # Saved Checklists tab
        saved_tab.create_saved_checklists_tab(
            self.tab_widgets["Saved Checklists"],
            load_checklist_callback=lambda path: self.load_checklist_file(path),
            notebook=self.tabs
        )

        # --- Footer buttons ---
        footer = QWidget()
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(6, 4, 6, 4)

        left_buttons = QHBoxLayout()
        left_buttons.setSpacing(8)
        btn_new = QPushButton("New Checklist")
        btn_new.clicked.connect(self.new_checklist_action)
        left_buttons.addWidget(btn_new)

        btn_save = QPushButton("Save Checklist")
        btn_save.clicked.connect(self.save_checklist_file)
        left_buttons.addWidget(btn_save)

        # Optional Save As:
        # btn_save_as = QPushButton("Save Checklist As")
        # btn_save_as.clicked.connect(self.save_checklist_file_as)
        # left_buttons.addWidget(btn_save_as)

        btn_export_html = QPushButton("Export HTML")
        btn_export_html.clicked.connect(self.export_html_action)
        left_buttons.addWidget(btn_export_html)

        footer_layout.addLayout(left_buttons)
        footer_layout.addStretch()

        right_container = QWidget()
        right_layout = QHBoxLayout(right_container)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setAlignment(Qt.AlignRight)

        release_icon = QLabel("📄")
        release_icon.setToolTip("Release Notes")
        release_icon.setCursor(QCursor(Qt.PointingHandCursor))
        release_icon.setStyleSheet("QLabel { font-size: 20px; }")
        release_icon.mousePressEvent = lambda event: self.show_release_notes()

        instructions_icon = QLabel("❓")
        instructions_icon.setToolTip("How to Use")
        instructions_icon.setCursor(QCursor(Qt.PointingHandCursor))
        instructions_icon.setStyleSheet("QLabel { font-size: 20px; }")
        instructions_icon.mousePressEvent = lambda event: self.show_instructions()

        right_layout.addWidget(release_icon)
        right_layout.addWidget(instructions_icon)
        footer_layout.addWidget(right_container)

        main_layout = QVBoxLayout()
        main_layout.addWidget(self.tabs)
        main_layout.addWidget(footer)
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # Keyboard shortcuts
        btn_new.setShortcut(QKeySequence("Ctrl+N"))
        btn_save.setShortcut(QKeySequence("Ctrl+S"))
        btn_export_html.setShortcut(QKeySequence("Ctrl+H"))

        # Restore last tab
        last_tab = self.settings.get("last_tab_index", 1)
        if 0 <= last_tab < len(self.tab_names):
            self.tabs.setCurrentIndex(last_tab)
        else:
            self.tabs.setCurrentIndex(1)

        # Override closeEvent
        self._original_closeEvent = self.closeEvent
        self.closeEvent = self.on_close


    def is_any_dirty(self):
        return any(dt.is_dirty() for dt in self.dirty_trackers.values())


    def ask_save_discard_cancel(self):
        if not self.is_any_dirty():
            return "continue"

        msg = QMessageBox(self)
        msg.setWindowTitle("Unsaved Changes")
        msg.setText("There are unsaved changes. What do you want to do?")
        btn_save = msg.addButton("Save and continue", QMessageBox.AcceptRole)
        btn_discard = msg.addButton("Discard changes", QMessageBox.DestructiveRole)
        btn_cancel = msg.addButton("Cancel", QMessageBox.RejectRole)
        msg.setDefaultButton(btn_save)
        msg.exec()

        clicked = msg.clickedButton()
        if clicked == btn_save:
            self.save_checklist_file()
            if self.is_any_dirty():
                return "cancel"
            return "continue"
        elif clicked == btn_discard:
            return "continue"
        else:
            return "cancel"


    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)


    # ─── New helper: validate mandatory top-fields ─────────────────────────────
    def validate_top_fields(self):
        """
        Ensure Customer Name, Opp Name, ID # and Sales/CSR are all filled.
        If any are blank, warn the user, switch to the Checklist tab, and return False.
        """
        missing = []
        labels = ["Customer Name", "Opp Name", "ID #", "Sales/CSR"]
        for idx, label in enumerate(labels):
            text = self.checklist_tab.top_fields[idx].text().strip()
            if not text:
                missing.append(label)

        if missing:
            QMessageBox.warning(
                self,
                "Missing Fields",
                "Please fill out:\n" + "\n".join(missing)
            )
            # Switch to the Checklist tab
            chk_idx = self.tabs.indexOf(self.tab_widgets["Checklist"])
            if chk_idx != -1:
                self.tabs.setCurrentIndex(chk_idx)
            return False

        return True
    # ─────────────────────────────────────────────────────────────────────────────


    def cleanup_lockfile(self):
        if self.current_lockfile and os.path.exists(self.current_lockfile):
            try:
                os.remove(self.current_lockfile)
            except Exception:
                pass
            self.current_lockfile = None


    def save_checklist_file(self):
        # 1) Mandatory top-fields
        if not self.validate_top_fields():
            return

        # 2) Gather quote‐info & require at least one drawing
        qi_data = qi_tab.get_quote_info_data(include_all=True)
        has_drawing = any(
            row.get("fields", [""])[0].strip() not in ["", "Drawing Number"]
            for row in qi_data
        )
        if not has_drawing:
            QMessageBox.warning(
                self,
                "Cannot Save",
                "At least one Drawing Number must be entered before saving."
            )
            return

        # 3) Bundle all data
        data = {
            "checklist":    cl_tab.get_checklist_data(),
            "quote_info":   qi_data,
            "vendor_quotes": vq_tab.get_vendor_quote_data(),
            "notes":        self.get_additional_notes()
        }

        # 4) Auto-save logic: build default filename from drawings
        filename_parts = []
        for row in qi_data:
            fields = row.get("fields", [])
            if not fields:
                continue
            drawing = fields[0].strip().upper()
            if not drawing:
                continue
            matched = find_latest_revision_files(drawing)
            if matched:
                latest = max(
                    matched.keys(),
                    key=lambda r: [int(p) for p in r.split(".") if p.isdigit()]
                )
                filename_parts.append(f"{drawing}_Rev{latest}")
            else:
                filename_parts.append(f"{drawing}_Rev0")

        unique_sorted = sorted(set(filename_parts))
        default_name = (
            " ".join(unique_sorted) + ".json"
            if unique_sorted else "Checklist.json"
        )

        save_path = self.current_checklist_path or os.path.join(CHECKLISTS_DIR, default_name)
        self.current_checklist_path = save_path
        self.update_window_title()

        # 5) Attempt to save
        try:
            user_settings.save_combined_data(save_path, data)
            QMessageBox.information(
                self,
                "Saved",
                f"Checklist auto-saved as:\n{os.path.basename(save_path)}"
            )
            for dt in self.dirty_trackers.values():
                dt.mark_clean()
        except Exception as e:
            # Fallback to Save As if write fails
            QMessageBox.warning(
                self,
                "Save Failed",
                f"Auto-save failed:\n{e}\n\nPlease choose a new file name."
            )
            base = os.path.basename(save_path)
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Checklist As",
                os.path.join(CHECKLISTS_DIR, base),
                "JSON Files (*.json)"
            )
            if file_path:
                enforced = os.path.join(CHECKLISTS_DIR, os.path.basename(file_path))
                try:
                    user_settings.save_combined_data(enforced, data)
                    QMessageBox.information(
                        self,
                        "Saved",
                        f"Checklist saved as:\n{os.path.basename(enforced)}"
                    )
                    self.current_checklist_path = enforced
                    self.update_window_title()
                    for dt in self.dirty_trackers.values():
                        dt.mark_clean()
                except Exception as err:
                    QMessageBox.critical(
                        self,
                        "Save Failed",
                        f"Final save attempt failed:\n{err}"
                    )



    def save_checklist_file_as(self):
        # 1) Require top-fields & at least one drawing
        if not self.validate_top_fields():
            return

        qi_data = qi_tab.get_quote_info_data(include_all=True)
        has_drawing = any(
            row.get("fields", [""])[0].strip() not in ["", "Drawing Number"]
            for row in qi_data
        )
        if not has_drawing:
            QMessageBox.warning(
                self,
                "Cannot Save",
                "At least one Drawing Number must be entered before saving."
            )
            return

        # 2) Bundle data
        data = {
            "checklist":     cl_tab.get_checklist_data(),
            "quote_info":    qi_data,
            "vendor_quotes": vq_tab.get_vendor_quote_data(),
            "notes":         self.get_additional_notes()
        }

        # 3) Prompt for file name
        default_name = "Checklist.json"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Checklist As",
            os.path.join(CHECKLISTS_DIR, default_name),
            "JSON Files (*.json)"
        )
        if not file_path:
            return  # user canceled

        # 4) Force into your checklist folder
        base = os.path.basename(file_path)
        enforced = os.path.join(CHECKLISTS_DIR, base)
        self.current_checklist_path = enforced
        self.update_window_title()

        # 5) Write it out
        try:
            user_settings.save_combined_data(enforced, data)
            QMessageBox.information(
                self,
                "Saved",
                f"Checklist saved as:\n{base}"
            )
            for dt in self.dirty_trackers.values():
                dt.mark_clean()
        except Exception as e:
            QMessageBox.critical(
                self,
                "Save Failed",
                f"Failed to save checklist:\n{e}"
            )



    def get_additional_notes(self):
        if hasattr(an_tab, "get_notes_text"):
            return an_tab.get_notes_text()
        return ""


    def show_release_notes(self):
        notes_path = r"P:\ENGINEERING\Design Checklist\supporting_documents\release_notes.txt"
        if not os.path.exists(notes_path):
            QMessageBox.warning(self, "File Not Found", "Could not locate file.")
            return
        try:
            with open(notes_path, "r", encoding="utf-8") as f:
                notes = f.read()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not read release notes:\n{e}")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Release Notes")
        dialog.setMinimumWidth(860)   # ⬅️ ensure popup never gets narrower than 860px
        dialog.resize(860, 500)       # optional: initial size
        layout = QVBoxLayout(dialog)
        text_box = QTextEdit()
        text_box.setReadOnly(True)
        text_box.setText(notes)
        layout.addWidget(text_box)
        dialog.exec()


    def show_instructions(self):
        instructions_path = r"P:\ENGINEERING\Design Checklist\supporting_documents\how_to.pdf"
        if not os.path.exists(instructions_path):
            QMessageBox.warning(self, "File Not Found", "Could not locate file.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(instructions_path))


    def load_checklist_file(self, path=None, force_read_only=False):
        if self.ask_save_discard_cancel() != "continue":
            return
        if not path:
            path, _ = QFileDialog.getOpenFileName(
                self,
                "Load Checklist",
                CHECKLISTS_DIR,
                "JSON Files (*.json)"
            )
        if not path or not os.path.exists(path):
            return

        self.current_checklist_path = path
        self.update_window_title()
        self.setUpdatesEnabled(False)

        try:
            # — Locking logic (unchanged) —
            self.cleanup_lockfile()
            locked, locker = lock_checklist(path)
            if not locked and not force_read_only:
                QMessageBox.warning(
                    self,
                    "Checklist In Use",
                    f"This checklist is currently being edited by: {locker or 'someone else'}.\n\n"
                    "You can open it in read-only mode."
                )
                force_read_only = True
            elif locked:
                self.current_lockfile = get_lock_path(path)

            # — Load combined data —
            data = user_settings.load_combined_data(path)
            if not data:
                return

            # — Backwards-compatible top_fields remap —
            top = data.get("checklist", {}).get("top_fields", [])
            if len(top) == 3:
                data["checklist"]["top_fields"] = [top[0], top[1], "", top[2]]
            elif len(top) < 3:
                data["checklist"]["top_fields"] = top + [""] * (4 - len(top))
            # else: len(top)>=4 → leave as-is

            # — Delegate into your tabs —
            cl_tab.load_checklist_data(data["checklist"], read_only=force_read_only)
            if "quote_info" in data:
                qi_tab.load_quote_info_data(data["quote_info"])
            if "vendor_quotes" in data:
                vq_tab.load_vendor_quote_data(data["vendor_quotes"])
            if "notes" in data and hasattr(an_tab, "set_notes_text"):
                an_tab.set_notes_text(data["notes"])
            elif hasattr(an_tab, "set_notes_text"):
                an_tab.set_notes_text("")

            # — Clear Rotary Reference —
            if hasattr(self, "gap_calc_widget") and hasattr(self.gap_calc_widget, "clear_fields"):
                self.gap_calc_widget.clear_fields()

        except Exception as e:
            QMessageBox.critical(self, "Load Failed", f"Could not load checklist:\n{e}")

        # mark clean & re-enable
        for dt in self.dirty_trackers.values():
            dt.mark_clean()
        self.setUpdatesEnabled(True)

        # switch to Checklist tab
        idx = self.tabs.indexOf(self.tab_widgets["Checklist"])
        if idx != -1:
            self.tabs.setCurrentIndex(idx)


    def new_checklist_action(self):
        self.current_checklist_path = None
        self.update_window_title()
        if self.ask_save_discard_cancel() != "continue":
            return
        self.cleanup_lockfile()
        cl_tab.clear_checklist_tab()
        qi_tab.clear_quote_info_tab()
        vq_tab.clear_vendor_quote_tab()
        if hasattr(saved_tab, "clear_saved_checklists_tab"):
            saved_tab.clear_saved_checklists_tab()
        if hasattr(an_tab, "set_notes_text"):
            an_tab.set_notes_text("")
        if hasattr(self, "gap_calc_widget") and self.gap_calc_widget:
            self.gap_calc_widget.clear_fields()
        self.tabs.setCurrentIndex(self.tab_names.index("Checklist"))
        for dt in self.dirty_trackers.values():
            dt.mark_clean()


    def export_html_action(self):
        qi_data = qi_tab.get_quote_info_data(include_all=True)
        vendor_quotes_data = vq_tab.get_vendor_quote_data()
        checklist_data = cl_tab.get_checklist_data()
        notes = self.get_additional_notes()

        drawings = set()
        for row in qi_data:
            fields = row.get("fields", [])
            if fields and fields[0].strip():
                drawings.add(fields[0].strip().upper())

        filename_parts = []
        for drawing in sorted(drawings):
            matched = find_latest_revision_files(drawing)
            if matched:
                latest = max(
                    matched.keys(),
                    key=lambda r: [int(p) for p in r.split(".") if p.isdigit()]
                )
                filename_parts.append(f"{drawing}_Rev{latest}")
            else:
                filename_parts.append(f"{drawing}_Rev0")

        unique_sorted = sorted(set(filename_parts))
        default_name = (
            " ".join(unique_sorted) + ".html"
            if unique_sorted else "checklist_export.html"
        )
        initial_dir = self.settings.get("last_html_export_dir", os.path.expanduser("~"))
        default_path = os.path.join(initial_dir, default_name)

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export HTML",
            default_path,
            "HTML Files (*.html)"
        )
        if not file_path:
            return

        self.settings["last_html_export_dir"] = os.path.dirname(file_path)
        user_settings.save_user_settings(self.settings)

        data = {
            "checklist": checklist_data,
            "quote_info": qi_data,
            "vendor_quotes": vendor_quotes_data,
            "notes": notes,
        }
        try:
            html_export.export_to_html(data, file_path)
            QMessageBox.information(self, "Export Complete", "HTML exported successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export HTML:\n{e}")


    def on_close(self, event):
        if self.is_any_dirty():
            action = self.ask_save_discard_cancel()
            if action == "cancel":
                event.ignore()
                return
        self.cleanup_lockfile()
        self.settings["window_geometry"] = f"{self.width()}x{self.height()}+{self.x()}+{self.y()}"
        self.settings["last_tab_index"] = self.tabs.currentIndex()
        user_settings.save_user_settings(self.settings)
        if hasattr(saved_tab, "save_column_widths_global"):
            saved_tab.save_column_widths_global()
        event.accept()


    def showEvent(self, event):
        super().showEvent(event)
        for dt in self.dirty_trackers.values():
            dt.mark_clean()


    def update_window_title(self):
        filename = (
            os.path.basename(self.current_checklist_path)
            if self.current_checklist_path else None
        )
        dirty = self.is_any_dirty()
        if filename and filename.lower().endswith(".json"):
            filename = filename[:-5]
        title = f"Engineering Checklist {APP_VERSION}"
        if filename:
            title += f" - {filename}"
        if dirty:
            title += " *"
        self.setWindowTitle(title)


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
