import json
import logging
import webbrowser

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTextEdit, QPushButton, QLabel,
    QComboBox, QApplication, QSlider, QSpinBox, QCheckBox, QListWidget,
    QListWidgetItem, QRadioButton, QAbstractItemView, QGroupBox, QMessageBox,
    QFrame, QFormLayout, QWidget, QScrollArea
)

from core.utils.paths import get_project_root



# Re-exported classes and constants from modularized dialogs
from modules.meetings.ui.tdoc_info_dialog import TDocInfoDialog
from modules.meetings.ui.excel_export_dialog import (
    ALL_AVAILABLE_COLUMNS,
    DEFAULT_SELECTED_COLUMNS,
    ExcelExportDialog,
)


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
        save_btn.setToolTip('Save My Status and My Notes for this TDoc.')
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

        self.config_path = get_project_root() / "config" / "stats_config.json"
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
        self.slider.setToolTip('Adjust community-detection sensitivity: lower values favor fewer/larger factions; higher values favor more/smaller factions.')
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
        self.thresh_spin.setToolTip('Minimum number of shared documents required before a contributor relationship is included in the graph.')
        self.thresh_spin.setRange(1, 20)
        self.thresh_spin.setValue(self.config.get("threshold", 1))
        self.thresh_spin.setStyleSheet("padding: 4px; border: 1px solid #CCC; background: white; width: 60px;")
        thresh_layout.addWidget(self.thresh_spin)
        layout.addLayout(thresh_layout)

        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("Top Contributors to Display in Chart:"))
        top_layout.addStretch()
        self.top_spin = QSpinBox()
        self.top_spin.setToolTip('Maximum number of top contributors displayed in the contributor chart.')
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
        self.hm_comp_spin.setToolTip('Maximum number of companies displayed in the company heatmap.')
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
        self.hm_ai_spin.setToolTip('Maximum number of Agenda Items displayed in the Agenda Item heatmap.')
        self.hm_ai_spin.setRange(5, 200)
        self.hm_ai_spin.setSingleStep(5)
        self.hm_ai_spin.setValue(self.config.get("heatmap_top_ais", 25))
        self.hm_ai_spin.setStyleSheet("padding: 4px; border: 1px solid #CCC; background: white; width: 60px;")
        hm_ai_layout.addWidget(self.hm_ai_spin)
        layout.addLayout(hm_ai_layout)

        self.export_html_chk = QCheckBox("Export standalone HTML plots (Warning: Creates hundreds of MBs of files)")
        self.export_html_chk.setToolTip('Also export standalone interactive HTML plots. This can create very large output files.')
        self.export_html_chk.setChecked(self.config.get("export_html_plots", False))
        self.export_html_chk.setStyleSheet("font-weight: bold; color: #D83B01; margin-top: 5px;")
        layout.addWidget(self.export_html_chk)

        layout.addStretch()

        # --- Buttons ---
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("💾 Save && Apply")
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
        # self.config["llm_max_chars"] = ... (removed legacy LLM config from statistics)
        # self.config["llm_system_prompt"] = ... (removed legacy LLM config from statistics)

        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            logging.error(f"[StatisticsSettings] Failed to save configuration: {e}", exc_info=True)

        self.accept()


