# --- File: modules/emails/ui/config_dialog.py ---
import platform
import logging
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QLineEdit, QPushButton, QTreeView, QMessageBox)
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

        # Intercept expansions to lazy-load subfolders
        self.tree.expanded.connect(self._load_children)

        self._init_root_folders()

    def _init_root_folders(self):
        if platform.system() != 'Windows':
            QMessageBox.warning(self, "Unsupported", "Outlook integration is only available on Windows.")
            return

        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()  # Ensure safe COM threading in PyQt

            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            for folder in namespace.Folders:
                item = QStandardItem(folder.Name)
                # Normalize the Outlook COM path (e.g., \\user@domain.com\Inbox -> user@domain.com/Inbox)
                clean_path = folder.FolderPath.replace('\\', '/').strip('/')
                item.setData(clean_path, Qt.UserRole)
                item.setData(folder, Qt.UserRole + 1)  # Store the COM object securely

                # Add a dummy child to trigger the expand arrow if subfolders exist
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
        # If the only child is our dummy "Loading..." node, replace it with real folders
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
    def __init__(self, current_source: str, current_target: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("⚙️ Configure Outlook Folders")
        self.resize(600, 200)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; }")

        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Source Folder UI
        layout.addWidget(QLabel("<b>Source Folder</b> (Where new eMeeting list emails arrive):"))
        src_layout = QHBoxLayout()
        self.txt_source = QLineEdit(current_source)
        self.txt_source.setPlaceholderText("e.g. user@domain.com/Inbox/3GPP List")
        btn_src_browse = QPushButton("Browse Outlook...")
        btn_src_browse.clicked.connect(lambda: self._browse_folder(self.txt_source))
        src_layout.addWidget(self.txt_source)
        src_layout.addWidget(btn_src_browse)
        layout.addLayout(src_layout)

        # Target Folder UI
        layout.addWidget(QLabel("<b>Target Folder</b> (Base folder where processed emails will be moved):"))
        tgt_layout = QHBoxLayout()
        self.txt_target = QLineEdit(current_target)
        self.txt_target.setPlaceholderText("e.g. user@domain.com/Archive/SA2_175")
        btn_tgt_browse = QPushButton("Browse Outlook...")
        btn_tgt_browse.clicked.connect(lambda: self._browse_folder(self.txt_target))
        tgt_layout.addWidget(self.txt_target)
        tgt_layout.addWidget(btn_tgt_browse)
        layout.addLayout(tgt_layout)

        layout.addStretch()

        # Save / Cancel Buttons
        btn_layout = QHBoxLayout()
        btn_save = QPushButton("Save Configuration")
        btn_save.setStyleSheet(
            "font-weight: bold; background-color: #0C6B0C; color: white; padding: 6px 15px; border-radius: 4px;")
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

    def get_paths(self):
        return self.txt_source.text().strip(), self.txt_target.text().strip()