# --- File: src/core/ui/process_manager_dialog.py ---
import logging

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor
from PyQt5.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMenu,
    QMessageBox,
    QPushButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
)

from core.process_manager import ProcessManager
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_DANGER,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    BUTTON_STYLE_TOOLBAR_WARNING,
    TREE_STYLE_CLEAN,
)

logger = logging.getLogger(__name__)

class ProcessManagerDialog(QDialog):
    """Compact Office document/process manager with safe document-level closing."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("COM Process Manager")
        self.setWindowFlags(Qt.Window)
        self.resize(980, 600)

        self.apps = {
            "Microsoft Word": "winword",
            "Microsoft PowerPoint": "powerpnt",
            "Microsoft Visio": "visio",
        }
        # Session-only set of documents that the user has unchecked / chosen to keep open.
        self._kept_open = {"winword": set(), "powerpnt": set()}
        self._document_items = []
        self._app_items = {}
        self._updating_checks = False

        self._setup_ui()
        self._refresh_stats()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 14)
        layout.setSpacing(8)

        title = QLabel("Active Office Documents")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #1E293B;")
        layout.addWidget(title)

        desc = QLabel(
            "Checked documents will be closed. Uncheck anything you want to keep open. "
            "Process termination is available under Emergency."
        )
        desc.setStyleSheet("color: #64748B; font-size: 11px;")
        layout.addWidget(desc)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Close / Application / Document", "Location / Status"])
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tree.header().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tree.setAlternatingRowColors(False)
        self.tree.setRootIsDecorated(True)
        self.tree.setUniformRowHeights(True)
        self.tree.setStyleSheet(TREE_STYLE_CLEAN)
        self.tree.itemChanged.connect(self._on_item_changed)
        layout.addWidget(self.tree, 1)

        # Selection helpers are intentionally lightweight: they affect documents only.
        selection_bar = QHBoxLayout()
        selection_bar.setSpacing(6)
        selection_label = QLabel("Documents:")
        selection_label.setStyleSheet("color: #64748B; font-size: 11px;")
        selection_bar.addWidget(selection_label)

        select_all_btn = QPushButton("Select all")
        select_all_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        select_all_btn.clicked.connect(lambda: self._set_all_documents_checked(True))
        selection_bar.addWidget(select_all_btn)

        select_none_btn = QPushButton("Select none")
        select_none_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        select_none_btn.clicked.connect(lambda: self._set_all_documents_checked(False))
        selection_bar.addWidget(select_none_btn)
        selection_bar.addStretch()
        layout.addLayout(selection_bar)

        footer = QHBoxLayout()
        footer.setSpacing(8)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #64748B; font-size: 11px;")
        footer.addWidget(self.status_label)
        footer.addStretch()

        self.close_docs_btn = QPushButton("Close Selected")
        self.close_docs_btn.setToolTip(
            "Close checked Word and PowerPoint documents normally. Office may prompt to save modified documents."
        )
        self.close_docs_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.close_docs_btn.clicked.connect(lambda: self._close_selected_documents(False))
        footer.addWidget(self.close_docs_btn)

        self.discard_btn = QPushButton("Close && Discard Changes")
        self.discard_btn.setToolTip(
            "Close checked Word and PowerPoint documents and deliberately discard unsaved edits."
        )
        self.discard_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_WARNING)
        self.discard_btn.clicked.connect(lambda: self._close_selected_documents(True))
        footer.addWidget(self.discard_btn)

        self.emergency_btn = QPushButton("Emergency ▾")
        self.emergency_btn.setToolTip("Force-terminate Office processes. Use only when graceful closing is not possible.")
        self.emergency_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_DANGER)
        self.emergency_menu = QMenu(self.emergency_btn)
        self.emergency_btn.setMenu(self.emergency_menu)
        footer.addWidget(self.emergency_btn)

        refresh_btn = QPushButton("Refresh")
        refresh_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        refresh_btn.clicked.connect(self._refresh_stats)
        footer.addWidget(refresh_btn)

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.clicked.connect(self.accept)
        footer.addWidget(close_btn)
        layout.addLayout(footer)

    @staticmethod
    def _doc_key(doc):
        return (doc.get("full_name") or doc.get("name") or "").casefold()

    def _refresh_stats(self):
        expanded_states = {
            self.tree.topLevelItem(i).data(0, Qt.UserRole): self.tree.topLevelItem(i).isExpanded()
            for i in range(self.tree.topLevelItemCount())
        }

        self._updating_checks = True
        self.tree.clear()
        self._document_items = []
        self._app_items = {}
        self.emergency_menu.clear()
        data = ProcessManager.get_process_stats()

        total_documents = 0
        total_checked = 0

        for display_name, exe_name in self.apps.items():
            app_procs = [p for p in data if p["Name"].lower() == exe_name.lower()]
            documents = ProcessManager.get_open_office_documents(exe_name) if exe_name in self._kept_open else []
            if not app_procs and not documents:
                continue

            total = len(app_procs)
            ghosts = sum(1 for p in app_procs if p["IsGhost"])
            app_item = QTreeWidgetItem(self.tree)
            app_item.setData(0, Qt.UserRole, exe_name)
            app_item.setText(0, display_name)
            app_item.setForeground(0, QBrush(QColor("#1E293B")))
            font = app_item.font(0)
            font.setBold(True)
            app_item.setFont(0, font)
            self._app_items[exe_name] = app_item

            if documents:
                status = f"{len(documents)} document{'s' if len(documents) != 1 else ''} · {total} process{'es' if total != 1 else ''}"
                if ghosts:
                    status += f" · {ghosts} background"
                app_item.setText(1, status)

                for doc in documents:
                    child = QTreeWidgetItem(app_item)
                    key = self._doc_key(doc)
                    child.setData(0, Qt.UserRole, (exe_name, key))
                    child.setFlags(child.flags() | Qt.ItemIsUserCheckable)
                    checked = key not in self._kept_open[exe_name]
                    child.setCheckState(0, Qt.Checked if checked else Qt.Unchecked)
                    name = doc.get("name", "Untitled Document")
                    if doc.get("modified") is True:
                        name += "  • modified"
                        child.setForeground(0, QBrush(QColor("#B45309")))
                    child.setText(0, name)
                    child.setText(1, doc.get("full_name") or "Unsaved document")
                    child.setToolTip(0, "Checked = close this document; unchecked = keep it open")
                    child.setToolTip(1, doc.get("full_name") or "Unsaved document")
                    self._document_items.append(child)
                    total_documents += 1
                    total_checked += int(checked)
            else:
                # Background-only applications stay compact. Process details remain available
                # by expanding the application row, but no action buttons clutter the tree.
                app_item.setText(1, f"{total} background process{'es' if total != 1 else ''}")
                app_item.setForeground(1, QBrush(QColor("#64748B")))
                for proc in app_procs:
                    child = QTreeWidgetItem(app_item)
                    title = proc.get("Title", "").strip()
                    child.setText(0, title or "Background instance")
                    child.setText(1, f"PID {proc['Id']}" + (" · headless" if proc["IsGhost"] else ""))
                    child.setForeground(0, QBrush(QColor("#64748B")))

            # Document-bearing apps are useful expanded; background-only apps default collapsed.
            default_expanded = bool(documents)
            app_item.setExpanded(expanded_states.get(exe_name, default_expanded))

            kill_all = self.emergency_menu.addAction(f"Kill all {display_name} processes")
            kill_all.triggered.connect(lambda _, app=exe_name: self._kill_all(app))
            if ghosts:
                kill_ghosts = self.emergency_menu.addAction(f"Kill {display_name} background processes ({ghosts})")
                kill_ghosts.triggered.connect(lambda _, app=exe_name: self._kill_ghosts(app))

        if self.emergency_menu.actions():
            self.emergency_menu.addSeparator()
            kill_all_office = self.emergency_menu.addAction("Kill all Office processes")
            kill_all_office.triggered.connect(self._kill_all_office)

        self._updating_checks = False
        self._update_selection_status(total_documents, total_checked)

    def _on_item_changed(self, item, column):
        if self._updating_checks or column != 0:
            return
        payload = item.data(0, Qt.UserRole)
        if not isinstance(payload, tuple) or len(payload) != 2:
            return
        app_name, key = payload
        if item.checkState(0) == Qt.Checked:
            self._kept_open[app_name].discard(key)
        else:
            self._kept_open[app_name].add(key)
        self._update_selection_status()

    def _set_all_documents_checked(self, checked):
        self._updating_checks = True
        state = Qt.Checked if checked else Qt.Unchecked
        for item in self._document_items:
            item.setCheckState(0, state)
            app_name, key = item.data(0, Qt.UserRole)
            if checked:
                self._kept_open[app_name].discard(key)
            else:
                self._kept_open[app_name].add(key)
        self._updating_checks = False
        self._update_selection_status()

    def _update_selection_status(self, total=None, checked=None):
        if total is None:
            total = len(self._document_items)
        if checked is None:
            checked = sum(item.checkState(0) == Qt.Checked for item in self._document_items)
        kept = total - checked
        if total:
            text = f"{checked} selected to close"
            if kept:
                text += f" · {kept} kept open"
        else:
            text = "No closable Word or PowerPoint documents"
        self.status_label.setText(text)
        self.close_docs_btn.setEnabled(checked > 0)
        self.discard_btn.setEnabled(checked > 0)

    def _close_selected_documents(self, discard_changes):
        selected_by_app = {"winword": 0, "powerpnt": 0}
        for item in self._document_items:
            if item.checkState(0) == Qt.Checked:
                app_name, _ = item.data(0, Qt.UserRole)
                selected_by_app[app_name] += 1
        selected_count = sum(selected_by_app.values())
        if not selected_count:
            return

        if discard_changes:
            answer = QMessageBox.warning(
                self,
                "Discard Unsaved Changes?",
                f"Close {selected_count} selected document(s) and permanently discard any unsaved changes?\n\n"
                "Unchecked documents will remain open.",
                QMessageBox.Yes | QMessageBox.Cancel,
                QMessageBox.Cancel,
            )
            if answer != QMessageBox.Yes:
                return

        failures = []
        closed = 0
        excluded = 0
        try:
            for app_name, count in selected_by_app.items():
                if count <= 0:
                    continue
                result = ProcessManager.close_office_documents(
                    app_name,
                    excluded_paths=list(self._kept_open[app_name]),
                    discard_changes=discard_changes,
                )
                closed += result.get("closed", 0)
                excluded += result.get("excluded", 0)
                failures.extend(result.get("failed", []))
        except Exception as exc:
            logging.exception("Graceful Office close failed")
            QMessageBox.critical(self, "Close Failed", str(exc))
        else:
            if failures:
                QMessageBox.warning(
                    self,
                    "Some Documents Could Not Be Closed",
                    f"Closed: {closed}\nKept open: {excluded}\nFailed: {len(failures)}",
                )
        self._refresh_stats()

    def _kill_ghosts(self, app_name):
        ProcessManager.kill_app_ghosts(app_name)
        self._refresh_stats()

    def _kill_all(self, app_name):
        ProcessManager.kill_app_all(app_name)
        self._refresh_stats()

    def _kill_all_office(self):
        answer = QMessageBox.warning(
            self,
            "Kill All Office Processes?",
            "This force-terminates Word, PowerPoint, and Visio processes. Unsaved work may be lost.",
            QMessageBox.Yes | QMessageBox.Cancel,
            QMessageBox.Cancel,
        )
        if answer != QMessageBox.Yes:
            return
        for exe_name in self.apps.values():
            ProcessManager.kill_app_all(exe_name)
        self._refresh_stats()
