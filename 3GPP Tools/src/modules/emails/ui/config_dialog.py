# --- File: src/modules/emails/ui/config_dialog.py ---
import platform
import logging
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QLineEdit, QPushButton, QTreeView, QMessageBox, QSpinBox, QFrame)
from PyQt5.QtGui import QStandardItemModel, QStandardItem


class OutlookFolderPickerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Outlook Folder")
        self.resize(450, 500)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; }")

        layout = QVBoxLayout(self)

        self.tree = QTreeView()
        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels(["Outlook Folders"])
        self.tree.setModel(self.model)
        self.tree.setHeaderHidden(True)
        self.tree.setEditTriggers(QTreeView.NoEditTriggers)
        self.tree.setStyleSheet("QTreeView { background-color: #FFFFFF; border: 1px solid #CCC; }")
        layout.addWidget(self.tree)

        btn_layout = QHBoxLayout()
        self.btn_ok = QPushButton("Select Folder")
        self.btn_ok.setStyleSheet(
            "font-weight: bold; background-color: #0078D7; color: white; padding: 5px 15px; border-radius: 4px;")
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setStyleSheet("padding: 5px 15px;")

        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_cancel)
        btn_layout.addWidget(self.btn_ok)
        layout.addLayout(btn_layout)

        self.btn_ok.clicked.connect(self.accept)
        self.btn_cancel.clicked.connect(self.reject)

        self.tree.expanded.connect(self._load_children)
        self._init_root_folders()

    def _init_root_folders(self):
        if platform.system() != 'Windows':
            QMessageBox.warning(self, "Unsupported", "Outlook integration is only available on Windows.")
            return

        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()

            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            for folder in namespace.Folders:
                item = QStandardItem(folder.Name)
                clean_path = folder.FolderPath.replace('\\', '/').strip('/')
                item.setData(clean_path, Qt.UserRole)
                item.setData(folder, Qt.UserRole + 1)

                try:
                    if folder.Folders.Count > 0:
                        item.appendRow(QStandardItem("Loading..."))
                except Exception:
                    pass
                self.model.appendRow(item)

        except Exception as e:
            logging.error(f"Failed to load Outlook folders: {e}")
            QMessageBox.critical(self, "Outlook Error", f"Could not connect to Outlook.\n{e}")

    def _load_children(self, index):
        item = self.model.itemFromIndex(index)
        if item.rowCount() == 1 and item.child(0).text() == "Loading...":
            item.removeRow(0)
            folder = item.data(Qt.UserRole + 1)
            if folder:
                try:
                    for subfolder in folder.Folders:
                        sub_item = QStandardItem(subfolder.Name)
                        clean_path = subfolder.FolderPath.replace('\\', '/').strip('/')
                        sub_item.setData(clean_path, Qt.UserRole)
                        sub_item.setData(subfolder, Qt.UserRole + 1)
                        try:
                            if subfolder.Folders.Count > 0:
                                sub_item.appendRow(QStandardItem("Loading..."))
                        except Exception:
                            pass
                        item.appendRow(sub_item)
                except Exception as e:
                    logging.warning(f"Could not read subfolders for {item.text()}: {e}")

    def get_selected_path(self):
        index = self.tree.currentIndex()
        if index.isValid():
            item = self.model.itemFromIndex(index)
            return item.data(Qt.UserRole)
        return ""


class EmailConfigDialog(QDialog):
    def __init__(self, current_source: str, current_target: str, current_stats_cfg: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("⚙️ Email Manager Configuration")
        self.resize(600, 450)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; }")

        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # --- FOLDER CONFIGURATION ---
        layout.addWidget(QLabel("<b>📂 Outlook Folder Configuration</b>"))

        src_layout = QHBoxLayout()
        self.txt_source = QLineEdit(current_source)
        self.txt_source.setPlaceholderText("e.g. user@domain.com/Inbox/3GPP List")
        btn_src_browse = QPushButton("Browse Outlook...")
        btn_src_browse.clicked.connect(lambda: self._browse_folder(self.txt_source))
        src_layout.addWidget(QLabel("Source:"))
        src_layout.addWidget(self.txt_source)
        src_layout.addWidget(btn_src_browse)
        layout.addLayout(src_layout)

        tgt_layout = QHBoxLayout()
        self.txt_target = QLineEdit(current_target)
        self.txt_target.setPlaceholderText("e.g. user@domain.com/Archive/SA2_175")
        btn_tgt_browse = QPushButton("Browse Outlook...")
        btn_tgt_browse.clicked.connect(lambda: self._browse_folder(self.txt_target))
        tgt_layout.addWidget(QLabel("Target: "))
        tgt_layout.addWidget(self.txt_target)
        tgt_layout.addWidget(btn_tgt_browse)
        layout.addLayout(tgt_layout)

        # Add a visual separator
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        layout.addWidget(line)

        # --- STATISTICS CONFIGURATION ---
        layout.addWidget(QLabel("<b>📊 Analytics Dashboard Configuration</b>"))

        def create_spinbox_row(label_text, default_val):
            row_layout = QHBoxLayout()
            row_layout.addWidget(QLabel(label_text))
            row_layout.addStretch()
            spin = QSpinBox()
            spin.setRange(5, 100)
            spin.setSingleStep(5)
            spin.setValue(default_val)
            spin.setStyleSheet("padding: 4px; border: 1px solid #CCC; background: white; width: 60px;")
            row_layout.addWidget(spin)
            layout.addLayout(row_layout)
            return spin

        self.spin_top_comps = create_spinbox_row("Top Active Companies to Display:",
                                                 current_stats_cfg.get("email_top_companies", 25))
        self.spin_top_dels = create_spinbox_row("Top Companies by Active Delegates:",
                                                current_stats_cfg.get("email_top_delegates", 25))
        self.spin_hm_comps = create_spinbox_row("Heatmap: Top Companies Matrix Limit:",
                                                current_stats_cfg.get("email_heatmap_top_comps", 25))
        self.spin_hm_ais = create_spinbox_row("Heatmap: Top Agenda Items Matrix Limit:",
                                              current_stats_cfg.get("email_heatmap_top_ais", 25))

        layout.addStretch()

        # --- BUTTONS ---
        btn_layout = QHBoxLayout()
        btn_save = QPushButton("Save Configuration")
        btn_save.setStyleSheet(
            "font-weight: bold; background-color: #0078D7; color: white; padding: 6px 15px; border-radius: 4px;")
        btn_cancel = QPushButton("Cancel")

        btn_save.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(btn_cancel)
        btn_layout.addWidget(btn_save)
        layout.addLayout(btn_layout)

    def _browse_folder(self, target_line_edit):
        dialog = OutlookFolderPickerDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            path = dialog.get_selected_path()
            if path:
                target_line_edit.setText(path)

    def get_config_data(self):
        return {
            "source_folder": self.txt_source.text().strip(),
            "target_folder": self.txt_target.text().strip(),
            "stats_config": {
                "email_top_companies": self.spin_top_comps.value(),
                "email_top_delegates": self.spin_top_dels.value(),
                "email_heatmap_top_comps": self.spin_hm_comps.value(),
                "email_heatmap_top_ais": self.spin_hm_ais.value()
            }
        }