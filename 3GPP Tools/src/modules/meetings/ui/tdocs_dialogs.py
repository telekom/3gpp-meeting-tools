# --- File: src/modules/meetings/ui/tdocs_dialogs.py ---
import json
from pathlib import Path
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTextEdit,
                             QPushButton, QLabel, QComboBox, QApplication,
                             QSlider, QSpinBox)
from PyQt5.QtCore import Qt


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
        self.status_combo.addItems(["⚪ Neutral", "🟢 Support", "🔴 Object", "🟡 Monitor"])
        self.status_combo.setStyleSheet("padding: 4px; border: 1px solid #CCC; background: white;")

        curr_status = row_data.get("My Status", "⚪ Neutral").replace("🔄 ", "").strip()
        self.status_combo.setCurrentText(
            curr_status if curr_status in ["⚪ Neutral", "🟢 Support", "🔴 Object", "🟡 Monitor"] else "⚪ Neutral")

        status_layout.addWidget(self.status_combo)
        status_layout.addStretch()
        layout.addLayout(status_layout)

        layout.addWidget(QLabel("<b>My Notes:</b>"))
        self.my_notes = QTextEdit()
        clean_notes = row_data.get("My Notes", "").replace("🔄 [From Base]: ", "").replace("🔄 [From Base]", "").strip()
        self.my_notes.setPlainText(clean_notes)
        self.my_notes.setStyleSheet(
            "font-size: 13px; padding: 10px; background-color: white; border: 1px solid #005A9E;")
        layout.addWidget(self.my_notes)

        btn_layout = QHBoxLayout()
        save_btn = QPushButton("💾 Save Notes")
        save_btn.setStyleSheet(
            "padding: 6px 15px; font-weight: bold; background-color: #0C6B0C; color: white; border-radius: 4px;")
        save_btn.clicked.connect(self._on_save_clicked)

        btn_layout.addStretch()
        btn_layout.addWidget(save_btn)
        layout.addLayout(btn_layout)

    def _on_save_clicked(self):
        self.db_save_callback(self.tdoc_id, self.status_combo.currentText(), self.my_notes.toPlainText())
        self.accept()


class StatisticsSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("⚙️ Statistics Configuration")
        self.resize(450, 250)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; } QLabel { font-size: 13px; color: #333; }")

        # Resolve the root src/ directory to ensure the config applies globally to all meetings
        self.config_path = Path(__file__).resolve().parents[4] / "stats_config.json"
        self.config = self.load_config()

        layout = QVBoxLayout(self)

        # --- Granularity Slider ---
        layout.addWidget(QLabel("<b>Faction Granularity (Algorithm Sensitivity)</b>"))
        desc_lbl = QLabel(
            "<i>Controls how the math groups co-signers. Slide left for a few massive alliances, slide right to detect many small, strict factions.</i>")
        desc_lbl.setWordWrap(True)
        desc_lbl.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(desc_lbl)

        self.slider = QSlider(Qt.Horizontal)
        self.slider.setMinimum(5)  # Represents 0.5 (Low resolution, large communities)
        self.slider.setMaximum(25)  # Represents 2.5 (High resolution, small communities)
        self.slider.setSingleStep(1)
        self.slider.setValue(int(self.config.get("resolution", 1.5) * 10))
        layout.addWidget(self.slider)

        slider_labels = QHBoxLayout()
        slider_labels.addWidget(QLabel("Fewer / Massive Factions"))
        slider_labels.addStretch()
        slider_labels.addWidget(QLabel("Many / Small Factions"))
        layout.addLayout(slider_labels)

        layout.addSpacing(15)

        # --- Threshold Spinbox ---
        thresh_layout = QHBoxLayout()
        thresh_layout.addWidget(QLabel("<b>Minimum Shared Documents (Graph Filter):</b>"))
        thresh_layout.addStretch()
        self.thresh_spin = QSpinBox()
        self.thresh_spin.setRange(1, 20)
        self.thresh_spin.setValue(self.config.get("threshold", 1))
        self.thresh_spin.setStyleSheet("padding: 4px; border: 1px solid #CCC; background: white; width: 60px;")
        thresh_layout.addWidget(self.thresh_spin)
        layout.addLayout(thresh_layout)

        # --- Top Contributors Spinbox ---
        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("<b>Top Contributors to Display in Chart:</b>"))
        top_layout.addStretch()
        self.top_spin = QSpinBox()
        self.top_spin.setRange(10, 100)
        self.top_spin.setSingleStep(5)
        self.top_spin.setValue(self.config.get("top_count", 30))
        self.top_spin.setStyleSheet("padding: 4px; border: 1px solid #CCC; background: white; width: 60px;")
        top_layout.addWidget(self.top_spin)
        layout.addLayout(top_layout)

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

    def load_config(self):
        default = {"resolution": 1.5, "threshold": 1, "top_count": 30}
        if self.config_path.exists():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return default

    def save_config(self):
        self.config["resolution"] = self.slider.value() / 10.0
        self.config["threshold"] = self.thresh_spin.value()
        self.config["top_count"] = self.top_spin.value()

        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            print(f"Failed to save statistics config: {e}")

        self.accept()