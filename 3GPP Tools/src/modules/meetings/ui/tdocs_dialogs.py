import inspect
import json
import logging
from typing import List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTextEdit, QPushButton, QLabel,
    QComboBox, QApplication, QSlider, QSpinBox, QCheckBox, QWidget
)

from core.ui.flow_layout import FlowLayout
from core.utils.paths import get_project_root
from modules.meetings.ui.topic_pills import TagCloudWidget



# Re-exported classes and constants from modularized dialogs
from modules.meetings.ui.tdoc_info_dialog import TDocInfoDialog
from modules.meetings.ui.dialogs.excel_export_dialog import (
    ALL_AVAILABLE_COLUMNS,
    DEFAULT_SELECTED_COLUMNS,
    ExcelExportDialog,
)
from core.ui.theme import C


class ReadOnlyViewerDialog(QDialog):
    def __init__(self, parent, title: str, text: str):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(600, 450)
        self.setStyleSheet(f"QDialog {{ background-color: {C.BG_ALT_LIGHT}; }}")

        layout = QVBoxLayout(self)
        self.text_edit = QTextEdit()
        self.text_edit.setPlainText(text)
        self.text_edit.setReadOnly(True)
        self.text_edit.setStyleSheet("font-size: 13px; padding: 10px; background-color: white; border: 1px solid #CCC;")
        layout.addWidget(self.text_edit)

        btn_layout = QHBoxLayout()
        copy_btn = QPushButton("📋 Copy All")
        copy_btn.setStyleSheet(
            f"padding: 6px 15px; font-weight: bold; background-color: {C.OFFICE_BLUE}; color: white; border-radius: 4px;")
        copy_btn.clicked.connect(lambda: [QApplication.clipboard().setText(text), self.accept()])

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(
            f"padding: 6px 15px; background-color: {C.BORDER_GRAY}; border: 1px solid #CCC; border-radius: 4px;")
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
    def __init__(
        self,
        parent,
        tdoc_id,
        row_data,
        db_save_callback,
        all_existing_tags: Optional[List[str]] = None,
    ):
        super().__init__(parent)
        self.setWindowTitle(f"📝 Notes, Status & Tags: {tdoc_id}")
        self.resize(660, 580)
        self.setStyleSheet(f"QDialog {{ background-color: {C.BG_ALT_LIGHT}; }}")
        self.db_save_callback = db_save_callback
        self.tdoc_id = tdoc_id
        self._all_existing_tags = list(all_existing_tags or [])

        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        is_thread = bool(row_data.get("_is_tracked_thread") or str(tdoc_id).upper().startswith("THR-"))
        context_lbl = "<b>Thread Summary:</b>" if is_thread else "<b>Secretary Remarks:</b>"
        layout.addWidget(QLabel(context_lbl))
        sec_remarks = QTextEdit()
        sec_remarks.setPlainText(row_data.get("Secretary Remarks", "") or row_data.get("Title", ""))
        sec_remarks.setReadOnly(True)
        sec_remarks.setMaximumHeight(90)
        sec_remarks.setStyleSheet(f"background-color: {C.BG_SURFACE}; border: 1px solid #CCC;")
        layout.addWidget(sec_remarks)

        status_layout = QHBoxLayout()
        status_layout.addWidget(QLabel("<b>My Status:</b>"))
        self.status_combo = QComboBox()
        self.status_combo.addItems(MY_STATUS_OPTIONS)
        self.status_combo.setStyleSheet("padding: 4px; border: 1px solid #CCC; background: white;")

        curr_status = str(row_data.get("My Status", "⚪ Neutral") or "⚪ Neutral").replace("🔄 ", "").strip()
        self.status_combo.setCurrentText(
            curr_status if curr_status in MY_STATUS_OPTIONS else "⚪ Neutral"
        )

        status_layout.addWidget(self.status_combo)
        status_layout.addStretch()
        layout.addLayout(status_layout)

        # Tags pill editor + clickable meeting tag suggestions (modeled on EditSpecAnnotationsDialog)
        layout.addWidget(QLabel("<b>My Tags:</b>"))
        self.tags_cloud = TagCloudWidget(
            placeholder="Type tag and press Enter or comma (e.g. Key Issue, Follow-up)...",
            pill_colors={"bg": C.ACCENT_SOFT, "fg": C.INFO_DARK, "border": C.INFO_BORDER_MEDIUM},
        )
        raw_tags = str(row_data.get("My Tags", "") or "").replace("🔄 [From Base]: ", "").replace("🔄 [From Base]", "").strip()
        self.tags_cloud.setText(raw_tags)
        self.tags_cloud.set_suggestions(self._all_existing_tags)
        self.tags_cloud.tagsChanged.connect(self._rebuild_tag_suggestions)
        layout.addWidget(self.tags_cloud)

        self._sugg_container = QWidget(self)
        self._sugg_layout = FlowLayout(self._sugg_container, margin=0, h_spacing=6, v_spacing=4)
        layout.addWidget(self._sugg_container)
        self._rebuild_tag_suggestions()

        layout.addWidget(QLabel("<b>My Notes:</b>"))
        self.my_notes = QTextEdit()
        clean_notes = str(row_data.get("My Notes", "") or "").replace("🔄 [From Base]: ", "").replace("🔄 [From Base]", "").strip()
        self.my_notes.setPlainText(clean_notes)
        self.my_notes.setStyleSheet(
            f"font-size: 13px; padding: 10px; background-color: white; border: 1px solid {C.OFFICE_BLUE};"
        )
        layout.addWidget(self.my_notes, 1)

        btn_layout = QHBoxLayout()
        save_btn = QPushButton("💾 Save Annotations")
        save_btn.setToolTip('Save My Status, My Tags, and My Notes.')
        save_btn.setStyleSheet(
            f"padding: 6px 15px; font-weight: bold; background-color: {C.SUCCESS}; color: white; border-radius: 4px;"
        )
        save_btn.clicked.connect(self._on_save_clicked)

        btn_layout.addStretch()
        btn_layout.addWidget(save_btn)
        layout.addLayout(btn_layout)

    def _rebuild_tag_suggestions(self):
        self._sugg_layout.clear()
        assigned_lower = {t.lower() for t in self.tags_cloud.tags()}
        available = [t for t in self._all_existing_tags if t.lower() not in assigned_lower]
        if not available:
            self._sugg_container.setVisible(False)
            return
        self._sugg_container.setVisible(True)
        lbl = QLabel("Suggested:")
        lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        self._sugg_layout.addWidget(lbl)
        for tag in available[:14]:
            btn = QPushButton(f"+ {tag}")
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(
                f"QPushButton {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_BODY}; border: 1px solid {C.BORDER_STRONG}; "
                f"border-radius: 10px; padding: 2px 8px; font-size: 11px; }} "
                f"QPushButton:hover {{ background-color: {C.ACCENT_SOFT}; color: {C.ACCENT_PRIMARY}; border-color: {C.INFO_BORDER_MEDIUM}; }}"
            )
            btn.clicked.connect(lambda _, t=tag: self.tags_cloud.add_tag(t))
            self._sugg_layout.addWidget(btn)

    def _on_save_clicked(self):
        status = self.status_combo.currentText()
        notes = self.my_notes.toPlainText()
        tags_str = self.tags_cloud.text()

        if self.db_save_callback:
            try:
                sig = inspect.signature(self.db_save_callback)
                params = list(sig.parameters.values())
                accepts_varargs = any(p.kind == inspect.Parameter.VAR_POSITIONAL for p in params)
                if accepts_varargs or len(params) >= 4:
                    self.db_save_callback(self.tdoc_id, status, notes, tags_str)
                else:
                    self.db_save_callback(self.tdoc_id, status, notes)
            except (ValueError, TypeError):
                self.db_save_callback(self.tdoc_id, status, notes)

        self.accept()


class StatisticsSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("⚙️ Global Tools Configuration")
        self.resize(550, 620)  # Made slightly taller to fit the new settings
        self.setStyleSheet(f"QDialog {{ background-color: {C.BG_ALT_LIGHT}; }} QLabel {{ font-size: 13px; color: #333; }}")

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
        self.export_html_chk.setStyleSheet(f"font-weight: bold; color: {C.PPT_ORANGE_PRIMARY}; margin-top: 5px;")
        layout.addWidget(self.export_html_chk)

        layout.addStretch()

        # --- Buttons ---
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("💾 Save && Apply")
        save_btn.setStyleSheet(
            f"padding: 6px 15px; font-weight: bold; background-color: {C.OFFICE_BLUE}; color: white; border-radius: 4px;")
        save_btn.clicked.connect(self.save_config)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet(
            f"padding: 6px 15px; background-color: {C.BORDER_GRAY}; border: 1px solid #CCC; border-radius: 4px;")
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


