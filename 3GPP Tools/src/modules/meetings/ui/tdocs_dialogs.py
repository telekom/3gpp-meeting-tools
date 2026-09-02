import json
import logging
import webbrowser
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTextEdit, QPushButton, QLabel,
    QComboBox, QApplication, QSlider, QSpinBox, QCheckBox, QListWidget,
    QListWidgetItem, QRadioButton, QAbstractItemView, QGroupBox, QMessageBox,
    QFrame, QFormLayout, QWidget, QScrollArea
)


class TDocInfoDialog(QDialog):
    """Modernized TDoc Details Dialog displaying all stored database attributes,

    secretary remarks, personal statuses, revisions, and FTP links.
    """

    def __init__(self, details: dict, docs_ftp_url: str = "", revisions: list = None, parent=None):
        super().__init__(parent)
        self.details = details
        self.docs_ftp_url = docs_ftp_url.rstrip('/') if docs_ftp_url else ""
        self.revisions = revisions or []

        tdoc_id = str(details.get("TDoc", "Unknown")).strip()
        tdoc_type = str(details.get("Type", "TDoc")).strip() or "TDoc"
        tdoc_status = str(details.get("TDoc Status", "")).strip()
        title = str(details.get("Title", "No Title Available")).strip()

        self.setWindowTitle(f"TDoc Details: {tdoc_id}")
        self.resize(650, 720)
        self.setMinimumWidth(580)
        self.setStyleSheet("""
            QDialog {
                background-color: #F8F9FA;
            }
            QFrame#cardFrame {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 8px;
            }
            QLabel {
                font-size: 13px;
                color: #2D3748;
            }
            QPushButton {
                padding: 6px 14px;
                font-size: 12px;
                border-radius: 4px;
                border: 1px solid #CBD5E0;
                background-color: #FFFFFF;
                color: #2D3748;
            }
            QPushButton:hover {
                background-color: #EDF2F7;
                border-color: #A0AEC0;
            }
            QPushButton#primaryActionBtn {
                background-color: #005A9E;
                color: #FFFFFF;
                border: 1px solid #004578;
                font-weight: bold;
            }
            QPushButton#primaryActionBtn:hover {
                background-color: #004578;
            }
            QPushButton#revChipBtn {
                background-color: #F0F4F8;
                border: 1px solid #D2E3FC;
                border-radius: 12px;
                padding: 3px 10px;
                font-size: 11px;
                color: #005A9E;
                font-weight: bold;
            }
            QPushButton#revChipBtn:hover {
                background-color: #E1EFFF;
                border-color: #005A9E;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # --- 1. HEADER CARD ---
        header_card = QFrame()
        header_card.setObjectName("cardFrame")
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(14, 12, 14, 12)
        header_layout.setSpacing(8)

        title_row = QHBoxLayout()
        title_row.setSpacing(8)

        # Document Type Badge
        type_badge = QLabel(f"<b>{tdoc_type}</b>")
        type_badge.setStyleSheet("""
            background-color: #EBF8FF;
            color: #2B6CB0;
            border: 1px solid #BEE3F8;
            border-radius: 4px;
            padding: 2px 8px;
            font-size: 12px;
            font-weight: bold;
        """)
        title_row.addWidget(type_badge)

        # Status Badge (if populated)
        if tdoc_status:
            status_badge = QLabel(f"<b>{tdoc_status}</b>")
            status_style = self._get_status_badge_style(tdoc_status)
            status_badge.setStyleSheet(status_style)
            title_row.addWidget(status_badge)

        tdoc_label = QLabel(f"<b>{tdoc_id}</b>")
        tdoc_label.setStyleSheet("font-size: 18px; color: #1A202C; font-weight: bold;")
        tdoc_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        title_row.addWidget(tdoc_label)

        title_row.addStretch()
        header_layout.addLayout(title_row)

        desc_label = QLabel(title)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #4A5568; font-size: 13px; line-height: 1.4;")
        desc_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        header_layout.addWidget(desc_label)

        layout.addWidget(header_card)

        # --- 2. SCROLLABLE DETAILS CARD ---
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: transparent; }")

        details_card = QFrame()
        details_card.setObjectName("cardFrame")
        form = QFormLayout(details_card)
        form.setContentsMargins(14, 14, 14, 14)
        form.setSpacing(10)
        form.setLabelAlignment(Qt.AlignRight)

        # Primary Metadata
        self._add_row(form, "Agenda Item", details.get("Agenda Item") or "-")
        self._add_row(form, "Source", details.get("Source") or "-")
        self._add_row(form, "Intended For", details.get("For") or "-")

        # Personal Status & Notes
        my_status = str(details.get("My Status", "")).replace("🔄 ", "").strip()
        if my_status and my_status != "⚪ Neutral":
            self._add_row(form, "My Status", my_status)

        my_notes = str(details.get("My Notes", "")).strip()
        if my_notes:
            self._add_row(form, "My Notes", my_notes)

        # Secretary Remarks
        sec_remarks = str(details.get("Secretary Remarks", "")).strip()
        if sec_remarks:
            self._add_row(form, "Secretary Remarks", sec_remarks)

        # Abstract
        abstract = str(details.get("Abstract", "")).strip()
        if abstract:
            self._add_row(form, "Abstract", abstract)

        # Related TDocs
        related = str(details.get("Related TDocs", "")).strip()
        if related:
            self._add_row(form, "Related TDocs", related)

        # Revision Lineage Chips
        if self.revisions:
            rev_container = QWidget()
            rev_layout = QHBoxLayout(rev_container)
            rev_layout.setContentsMargins(0, 0, 0, 0)
            rev_layout.setSpacing(6)

            for rev_id in self.revisions:
                chip = QPushButton(rev_id)
                chip.setObjectName("revChipBtn")
                chip.setCursor(Qt.PointingHandCursor)
                chip.setToolTip(f"Revision: {rev_id}")
                if self.docs_ftp_url:
                    chip.clicked.connect(lambda _, r=rev_id: webbrowser.open(f"{self.docs_ftp_url}/{r}.zip"))
                rev_layout.addWidget(chip)

            rev_layout.addStretch()
            form.addRow(self._make_key_label("Revisions:"), rev_container)

        # FTP Download URL
        if self.docs_ftp_url and tdoc_id != "Unknown":
            doc_url = f"{self.docs_ftp_url}/{tdoc_id}.zip"
            ftp_label = QLabel(f'<a href="{doc_url}" style="color: #005A9E; text-decoration: none;">{doc_url}</a>')
            ftp_label.setOpenExternalLinks(True)
            ftp_label.setTextInteractionFlags(Qt.TextBrowserInteraction | Qt.TextSelectableByMouse)
            form.addRow(self._make_key_label("FTP Download:"), ftp_label)

        # Catch-all: render any remaining attributes not explicitly covered
        excluded_keys = {
            "TDoc", "Type", "Title", "Source", "For", "Agenda Item",
            "TDoc Status", "Secretary Remarks", "My Status", "My Notes",
            "Abstract", "Related TDocs", "Emails"
        }
        for key, value in details.items():
            if key not in excluded_keys and not key.startswith("_") and value:
                display_key = key.replace("_", " ").title()
                self._add_row(form, display_key, str(value))

        scroll.setWidget(details_card)
        layout.addWidget(scroll)

        # --- 3. ACTION BUTTONS ---
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)

        copy_btn = QPushButton("📋 Copy Info")
        copy_btn.setCursor(Qt.PointingHandCursor)
        copy_btn.setToolTip("Copy formatted summary to clipboard")
        copy_btn.clicked.connect(self._copy_tdoc_summary)
        btn_layout.addWidget(copy_btn)

        if self.docs_ftp_url and tdoc_id != "Unknown":
            ftp_btn = QPushButton("📂 Open FTP Archive")
            ftp_btn.setObjectName("primaryActionBtn")
            ftp_btn.setCursor(Qt.PointingHandCursor)
            ftp_btn.clicked.connect(lambda: webbrowser.open(f"{self.docs_ftp_url}/{tdoc_id}.zip"))
            btn_layout.addWidget(ftp_btn)

        btn_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def _make_key_label(self, text: str) -> QLabel:
        lbl = QLabel(f"<b>{text}</b>")
        lbl.setStyleSheet("color: #718096; font-size: 12px;")
        return lbl

    def _add_row(self, form: QFormLayout, label_text: str, value_text: str):
        val_label = QLabel(value_text)
        val_label.setWordWrap(True)
        val_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        form.addRow(self._make_key_label(f"{label_text}:"), val_label)

    def _get_status_badge_style(self, status: str) -> str:
        s = status.lower()
        if "agreed" in s or "approved" in s:
            return "background-color: #E6F4EA; color: #137333; border: 1px solid #CEEAD6; border-radius: 4px; padding: 2px 8px; font-size: 12px; font-weight: bold;"
        if "noted" in s or "postponed" in s:
            return "background-color: #FEF7E0; color: #B06000; border: 1px solid #FEEFC3; border-radius: 4px; padding: 2px 8px; font-size: 12px; font-weight: bold;"
        if "revised" in s or "withdrawn" in s:
            return "background-color: #FCE8E6; color: #C5221F; border: 1px solid #FAD2CF; border-radius: 4px; padding: 2px 8px; font-size: 12px; font-weight: bold;"
        return "background-color: #F1F3F4; color: #3C4043; border: 1px solid #DADCE0; border-radius: 4px; padding: 2px 8px; font-size: 12px; font-weight: bold;"

    def _copy_tdoc_summary(self):
        lines = [
            f"TDoc: {self.details.get('TDoc', '')}",
            f"Title: {self.details.get('Title', '')}",
            f"Source: {self.details.get('Source', '')}",
            f"Type: {self.details.get('Type', '')}",
            f"Agenda Item: {self.details.get('Agenda Item', '')}",
            f"Status: {self.details.get('TDoc Status', '')}",
            f"For: {self.details.get('For', '')}",
        ]
        if self.details.get("Secretary Remarks"):
            lines.append(f"Remarks: {self.details.get('Secretary Remarks')}")
        if self.details.get("My Status"):
            lines.append(f"My Status: {self.details.get('My Status')}")
        if self.details.get("My Notes"):
            lines.append(f"My Notes: {self.details.get('My Notes')}")

        QApplication.clipboard().setText("\n".join(lines))
        QMessageBox.information(self, "Copied", "TDoc details copied to clipboard.")

class ReadOnlyViewerDialog(QDialog):
    def __init__(self, parent, title: str, text: str):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(600, 450)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; }")

        layout = QVBoxLayout(self)
        self.text_edit = QTextEdit()
        self.text_edit.setPlainText(text)
        self.text_edit.setReadOnly(True)
        self.text_edit.setStyleSheet("font-size: 13px; padding: 10px; background-color: white; border: 1px solid #CCC;")
        layout.addWidget(self.text_edit)

        btn_layout = QHBoxLayout()
        copy_btn = QPushButton("📋 Copy All")
        copy_btn.setStyleSheet(
            "padding: 6px 15px; font-weight: bold; background-color: #005A9E; color: white; border-radius: 4px;")
        copy_btn.clicked.connect(lambda: [QApplication.clipboard().setText(text), self.accept()])

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(
            "padding: 6px 15px; background-color: #E0E0E0; border: 1px solid #CCC; border-radius: 4px;")
        close_btn.clicked.connect(self.accept)

        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        btn_layout.addWidget(copy_btn)
        layout.addLayout(btn_layout)

# Define the master status list
MY_STATUS_OPTIONS = [
    "⚪ Neutral",
    "🔵 My TDoc",
    "🟢 Support",
    "🔴 Object",
    "🟡 Monitor"
]

class InteractiveNotesDialog(QDialog):
    def __init__(self, parent, tdoc_id, row_data, db_save_callback):
        super().__init__(parent)
        self.setWindowTitle(f"📝 Notes & Status: {tdoc_id}")
        self.resize(600, 500)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; }")
        self.db_save_callback = db_save_callback
        self.tdoc_id = tdoc_id

        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("<b>Secretary Remarks:</b>"))
        sec_remarks = QTextEdit()
        sec_remarks.setPlainText(row_data.get("Secretary Remarks", ""))
        sec_remarks.setReadOnly(True)
        sec_remarks.setMaximumHeight(100)
        sec_remarks.setStyleSheet("background-color: #F5F5F5; border: 1px solid #CCC;")
        layout.addWidget(sec_remarks)

        status_layout = QHBoxLayout()
        status_layout.addWidget(QLabel("<b>My Status:</b>"))
        self.status_combo = QComboBox()
        self.status_combo.addItems(MY_STATUS_OPTIONS)
        self.status_combo.setStyleSheet("padding: 4px; border: 1px solid #CCC; background: white;")

        curr_status = row_data.get("My Status", "⚪ Neutral").replace("🔄 ", "").strip()
        self.status_combo.setCurrentText(
            curr_status if curr_status in MY_STATUS_OPTIONS else "⚪ Neutral"
        )

        status_layout.addWidget(self.status_combo)
        status_layout.addStretch()
        layout.addLayout(status_layout)

        layout.addWidget(QLabel("<b>My Notes:</b>"))
        self.my_notes = QTextEdit()
        clean_notes = row_data.get("My Notes", "").replace("🔄 [From Base]: ", "").replace("🔄 [From Base]", "").strip()
        self.my_notes.setPlainText(clean_notes)
        self.my_notes.setStyleSheet(
            "font-size: 13px; padding: 10px; background-color: white; border: 1px solid #005A9E;"
        )
        layout.addWidget(self.my_notes)

        btn_layout = QHBoxLayout()
        save_btn = QPushButton("💾 Save Notes")
        save_btn.setStyleSheet(
            "padding: 6px 15px; font-weight: bold; background-color: #0C6B0C; color: white; border-radius: 4px;"
        )
        save_btn.clicked.connect(self._on_save_clicked)

        btn_layout.addStretch()
        btn_layout.addWidget(save_btn)
        layout.addLayout(btn_layout)

    def _on_save_clicked(self):
        status = self.status_combo.currentText()
        notes = self.my_notes.toPlainText()

        if self.db_save_callback:
            self.db_save_callback(self.tdoc_id, status, notes)

        self.accept()


class StatisticsSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("⚙️ Global Tools Configuration")
        self.resize(550, 620)  # Made slightly taller to fit the new settings
        self.setStyleSheet("QDialog { background-color: #FAFAFA; } QLabel { font-size: 13px; color: #333; }")

        self.config_path = Path(__file__).resolve().parents[4] / "stats_config.json"
        self.config = self.load_config()

        layout = QVBoxLayout(self)

        # --- Statistics Configuration ---
        layout.addWidget(QLabel("<b>📊 Faction Granularity (Algorithm Sensitivity)</b>"))
        desc_lbl = QLabel(
            "<i>Controls how the math groups co-signers. Slide left for a few massive alliances, slide right to detect many small, strict factions.</i>")
        desc_lbl.setWordWrap(True)
        desc_lbl.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(desc_lbl)

        self.slider = QSlider(Qt.Horizontal)
        self.slider.setMinimum(5)
        self.slider.setMaximum(25)
        self.slider.setSingleStep(1)
        self.slider.setValue(int(self.config.get("resolution", 1.5) * 10))
        layout.addWidget(self.slider)

        slider_labels = QHBoxLayout()
        slider_labels.addWidget(QLabel("Fewer / Massive Factions"))
        slider_labels.addStretch()
        slider_labels.addWidget(QLabel("Many / Small Factions"))
        layout.addLayout(slider_labels)
        layout.addSpacing(10)

        thresh_layout = QHBoxLayout()
        thresh_layout.addWidget(QLabel("Minimum Shared Documents (Graph Filter):"))
        thresh_layout.addStretch()
        self.thresh_spin = QSpinBox()
        self.thresh_spin.setRange(1, 20)
        self.thresh_spin.setValue(self.config.get("threshold", 1))
        self.thresh_spin.setStyleSheet("padding: 4px; border: 1px solid #CCC; background: white; width: 60px;")
        thresh_layout.addWidget(self.thresh_spin)
        layout.addLayout(thresh_layout)

        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("Top Contributors to Display in Chart:"))
        top_layout.addStretch()
        self.top_spin = QSpinBox()
        self.top_spin.setRange(10, 100)
        self.top_spin.setSingleStep(5)
        self.top_spin.setValue(self.config.get("top_count", 30))
        self.top_spin.setStyleSheet("padding: 4px; border: 1px solid #CCC; background: white; width: 60px;")
        top_layout.addWidget(self.top_spin)
        layout.addLayout(top_layout)

        # ---> NEW: Heatmap Dynamic Configurations
        hm_comp_layout = QHBoxLayout()
        hm_comp_layout.addWidget(QLabel("Heatmap: Top Companies to Display:"))
        hm_comp_layout.addStretch()
        self.hm_comp_spin = QSpinBox()
        self.hm_comp_spin.setRange(5, 200)
        self.hm_comp_spin.setSingleStep(5)
        self.hm_comp_spin.setValue(self.config.get("heatmap_top_companies", 25))
        self.hm_comp_spin.setStyleSheet("padding: 4px; border: 1px solid #CCC; background: white; width: 60px;")
        hm_comp_layout.addWidget(self.hm_comp_spin)
        layout.addLayout(hm_comp_layout)

        hm_ai_layout = QHBoxLayout()
        hm_ai_layout.addWidget(QLabel("Heatmap: Top Agenda Items to Display:"))
        hm_ai_layout.addStretch()
        self.hm_ai_spin = QSpinBox()
        self.hm_ai_spin.setRange(5, 200)
        self.hm_ai_spin.setSingleStep(5)
        self.hm_ai_spin.setValue(self.config.get("heatmap_top_ais", 25))
        self.hm_ai_spin.setStyleSheet("padding: 4px; border: 1px solid #CCC; background: white; width: 60px;")
        hm_ai_layout.addWidget(self.hm_ai_spin)
        layout.addLayout(hm_ai_layout)

        self.export_html_chk = QCheckBox("Export standalone HTML plots (Warning: Creates hundreds of MBs of files)")
        self.export_html_chk.setChecked(self.config.get("export_html_plots", False))
        self.export_html_chk.setStyleSheet("font-weight: bold; color: #D83B01; margin-top: 5px;")
        layout.addWidget(self.export_html_chk)

        layout.addSpacing(20)

        # --- LLM Configuration ---
        layout.addWidget(QLabel("<b>🤖 LLM Export Configuration</b>"))

        llm_desc = QLabel(
            "<i>Large meeting corpora will be split into multiple chunked files to prevent overflowing the AI's context limits. Customize the System Prompt to guide the LLM's analysis.</i>")
        llm_desc.setWordWrap(True)
        llm_desc.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(llm_desc)

        llm_layout = QHBoxLayout()
        llm_layout.addWidget(QLabel("Max Characters per File (Chunk Limit):"))
        llm_layout.addStretch()
        self.llm_spin = QSpinBox()
        self.llm_spin.setRange(10000, 5000000)
        self.llm_spin.setSingleStep(10000)
        self.llm_spin.setValue(self.config.get("llm_max_chars", 200000))
        self.llm_spin.setStyleSheet("padding: 4px; border: 1px solid #CCC; background: white; width: 80px;")
        llm_layout.addWidget(self.llm_spin)
        layout.addLayout(llm_layout)

        layout.addSpacing(10)
        layout.addWidget(QLabel("<b>System Prompt / Context Guide:</b>"))
        self.prompt_edit = QTextEdit()
        self.prompt_edit.setPlainText(self.config.get("llm_system_prompt", self._get_default_prompt()))
        self.prompt_edit.setStyleSheet(
            "padding: 8px; border: 1px solid #CCC; background: white; font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px;")
        layout.addWidget(self.prompt_edit)

        layout.addStretch()

        # --- Buttons ---
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("💾 Save & Apply")
        save_btn.setStyleSheet(
            "padding: 6px 15px; font-weight: bold; background-color: #005A9E; color: white; border-radius: 4px;")
        save_btn.clicked.connect(self.save_config)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet(
            "padding: 6px 15px; background-color: #E0E0E0; border: 1px solid #CCC; border-radius: 4px;")
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(save_btn)
        layout.addLayout(btn_layout)

    def _get_default_prompt(self):
        return (
            "This file contains a programmatic compilation of 3GPP Technical Documents (TDocs). "
            "These documents represent telecommunications standards proposals, revisions, and working group agreements.\n\n"
            "**Structural Rules for parsing this text:**\n"
            "- `[ADDED BLOCK]:` Denotes entirely new text inserted into the specification where tracking wasn't explicitly isolated.\n"
            "- `[INSERTED: <text>]`: Denotes specific inline text additions explicitly marked via Word Track Changes.\n"
            "- `[DELETED: <text>]`: Denotes specific inline text removals explicitly marked via Word Track Changes.\n\n"
            "**Your Task:** Please use this corpus to analyze technical agreements, architectural changes, or contradictions within this specific Agenda Item."
        )

    def load_config(self):
        default = {
            "resolution": 1.5,
            "threshold": 1,
            "top_count": 30,
            "heatmap_top_companies": 25,  # <--- NEW DEFAULT
            "heatmap_top_ais": 25,  # <--- NEW DEFAULT
            "export_html_plots": False,
            "llm_max_chars": 200000,
            "llm_system_prompt": self._get_default_prompt()
        }
        if self.config_path.exists():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    default.update(data)
                    return default
            except Exception:
                pass
        return default

    def save_config(self):
        self.config["resolution"] = self.slider.value() / 10.0
        self.config["threshold"] = self.thresh_spin.value()
        self.config["top_count"] = self.top_spin.value()
        self.config["heatmap_top_companies"] = self.hm_comp_spin.value()  # <--- SAVES VALUE
        self.config["heatmap_top_ais"] = self.hm_ai_spin.value()  # <--- SAVES VALUE
        self.config["export_html_plots"] = self.export_html_chk.isChecked()
        self.config["llm_max_chars"] = self.llm_spin.value()
        self.config["llm_system_prompt"] = self.prompt_edit.toPlainText().strip()

        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            print(f"Failed to save configuration: {e}")

        self.accept()


ALL_AVAILABLE_COLUMNS = [
    "TDoc",
    "Title",
    "Source",
    "Type",
    "Agenda Item",
    "TDoc Status",
    "For",
    "Secretary Remarks",
    "My Status",
    "My Notes",
    "Abstract",
    "Related TDocs"
]

DEFAULT_SELECTED_COLUMNS = [
    "TDoc",
    "Title",
    "Source",
    "Type",
    "Agenda Item",
    "TDoc Status"
]


class ExcelExportDialog(QDialog):
    def __init__(self, parent, visible_count: int, total_count: int):
        super().__init__(parent)
        self.setWindowTitle("📊 Export TDocs to Excel")
        self.resize(520, 600)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; } QLabel { font-size: 12px; color: #333; }")

        self.visible_count = visible_count
        self.total_count = total_count

        self.config_path = Path(__file__).resolve().parents[4] / "export_config.json"
        self.saved_config = self._load_saved_config()

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # --- Scope Selection ---
        scope_group = QGroupBox("1. Select Export Scope")
        scope_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        scope_layout = QVBoxLayout(scope_group)

        self.radio_visible = QRadioButton(f"Visible / Filtered Rows only ({self.visible_count} TDocs)")
        self.radio_all = QRadioButton(f"All Meeting Rows ({self.total_count} TDocs)")

        self.radio_visible.setChecked(True if self.visible_count > 0 else False)
        self.radio_all.setChecked(True if self.visible_count == 0 else False)

        scope_layout.addWidget(self.radio_visible)
        scope_layout.addWidget(self.radio_all)
        layout.addWidget(scope_group)

        # --- Columns Selection & Reordering ---
        col_group = QGroupBox("2. Choose Columns & Drag to Reorder")
        col_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        col_main_layout = QHBoxLayout(col_group)

        self.col_list = QListWidget()
        self.col_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.col_list.setDefaultDropAction(Qt.MoveAction)
        self.col_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.col_list.setStyleSheet("""
            QListWidget { 
                background: white; border: 1px solid #CCC; border-radius: 4px; padding: 4px; font-size: 12px; 
            }
            QListWidget::item { padding: 4px; }
            QListWidget::item:hover { background-color: #F0F4F8; }
        """)
        self._populate_column_list()
        col_main_layout.addWidget(self.col_list)

        # Action buttons for reordering & presets
        btn_box = QVBoxLayout()
        btn_box.setSpacing(6)

        btn_style = "QPushButton { padding: 5px 10px; font-size: 11px; background: white; border: 1px solid #CCC; border-radius: 4px; } QPushButton:hover { background: #EAEAEA; }"

        btn_up = QPushButton("▲ Move Up")
        btn_up.setStyleSheet(btn_style)
        btn_up.clicked.connect(self._move_item_up)

        btn_down = QPushButton("▼ Move Down")
        btn_down.setStyleSheet(btn_style)
        btn_down.clicked.connect(self._move_item_down)

        btn_all = QPushButton("☑️ Select All")
        btn_all.setStyleSheet(btn_style)
        btn_all.clicked.connect(lambda: self._set_all_checked(True))

        btn_none = QPushButton("◻️ Clear All")
        btn_none.setStyleSheet(btn_style)
        btn_none.clicked.connect(lambda: self._set_all_checked(False))

        btn_reset = QPushButton("🔄 Reset")
        btn_reset.setStyleSheet(btn_style)
        btn_reset.clicked.connect(self._reset_defaults)

        btn_box.addWidget(btn_up)
        btn_box.addWidget(btn_down)
        btn_box.addSpacing(10)
        btn_box.addWidget(btn_all)
        btn_box.addWidget(btn_none)
        btn_box.addWidget(btn_reset)
        btn_box.addStretch()
        col_main_layout.addLayout(btn_box)

        layout.addWidget(col_group)

        # --- Options ---
        self.chk_auto_open = QCheckBox("Automatically open spreadsheet after export")
        self.chk_auto_open.setChecked(self.saved_config.get("auto_open", True))
        self.chk_auto_open.setStyleSheet("font-size: 12px;")
        layout.addWidget(self.chk_auto_open)

        # --- Dialog Buttons ---
        btn_layout = QHBoxLayout()
        export_btn = QPushButton("📊 Export to Excel")
        export_btn.setStyleSheet("padding: 7px 20px; font-weight: bold; background-color: #005A9E; color: white; border-radius: 4px;")
        export_btn.clicked.connect(self._on_export_clicked)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet("padding: 7px 15px; background-color: #E0E0E0; border: 1px solid #CCC; border-radius: 4px;")
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(export_btn)
        layout.addLayout(btn_layout)

    def _populate_column_list(self):
        self.col_list.clear()
        saved_order = self.saved_config.get("column_order", ALL_AVAILABLE_COLUMNS)
        saved_checked = set(self.saved_config.get("checked_columns", DEFAULT_SELECTED_COLUMNS))

        # Add all saved columns preserving order
        ordered_cols = [c for c in saved_order if c in ALL_AVAILABLE_COLUMNS]
        # Append any missing columns that might have been added later
        for c in ALL_AVAILABLE_COLUMNS:
            if c not in ordered_cols:
                ordered_cols.append(c)

        for col_name in ordered_cols:
            item = QListWidgetItem(col_name)
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable | Qt.ItemIsDragEnabled)
            item.setCheckState(Qt.Checked if col_name in saved_checked else Qt.Unchecked)
            self.col_list.addItem(item)

    def _move_item_up(self):
        row = self.col_list.currentRow()
        if row > 0:
            item = self.col_list.takeItem(row)
            self.col_list.insertItem(row - 1, item)
            self.col_list.setCurrentRow(row - 1)

    def _move_item_down(self):
        row = self.col_list.currentRow()
        if row >= 0 and row < self.col_list.count() - 1:
            item = self.col_list.takeItem(row)
            self.col_list.insertItem(row + 1, item)
            self.col_list.setCurrentRow(row + 1)

    def _set_all_checked(self, checked: bool):
        state = Qt.Checked if checked else Qt.Unchecked
        for i in range(self.col_list.count()):
            self.col_list.item(i).setCheckState(state)

    def _reset_defaults(self):
        self.saved_config = {"column_order": ALL_AVAILABLE_COLUMNS, "checked_columns": DEFAULT_SELECTED_COLUMNS}
        self._populate_column_list()

    def _load_saved_config(self):
        if self.config_path.exists():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    return json.load(f).get("excel_export", {})
            except Exception:
                pass
        return {}

    def _save_current_config(self):
        order = []
        checked = []
        for i in range(self.col_list.count()):
            item = self.col_list.item(i)
            order.append(item.text())
            if item.checkState() == Qt.Checked:
                checked.append(item.text())

        full_config = {}
        if self.config_path.exists():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    full_config = json.load(f)
            except Exception:
                pass

        full_config["excel_export"] = {
            "column_order": order,
            "checked_columns": checked,
            "auto_open": self.chk_auto_open.isChecked()
        }

        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(full_config, f, indent=4)
        except Exception as e:
            logging.error(f"Failed to persist export configuration: {e}")

    def get_selected_columns(self):
        selected = []
        for i in range(self.col_list.count()):
            item = self.col_list.item(i)
            if item.checkState() == Qt.Checked:
                selected.append(item.text())
        return selected

    def is_visible_only(self):
        return self.radio_visible.isChecked()

    def should_auto_open(self):
        return self.chk_auto_open.isChecked()

    def _on_export_clicked(self):
        if not self.get_selected_columns():
            QMessageBox.warning(self, "No Columns Selected", "Please check at least one column to export.")
            return
        self._save_current_config()
        self.accept()