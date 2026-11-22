"""Visual Diff & AI Prompt UI setup mixin for SpecClauseDiffDialog."""

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QButtonGroup, QComboBox, QFrame, QGroupBox, QHBoxLayout, QLabel,
    QPushButton, QRadioButton, QTabWidget,
    QTextEdit, QVBoxLayout, QWidget,
)

from core.ui.theme import C


class _DiffUiMixin:
    """Builds Visual Diff tab, AI Prompt tab, action footer, and tab host."""

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        # ---------------------------------------------------------------------
        # Top Controls: Base and Target Version Selection
        # ---------------------------------------------------------------------
        config_group = QGroupBox("Release Comparison")
        config_layout = QHBoxLayout(config_group)
        config_layout.setContentsMargins(12, 8, 12, 8)
        config_layout.setSpacing(10)

        config_layout.addWidget(QLabel("<b>Base Release (Old):</b>"))
        self.base_ver_combo = QComboBox()
        self.base_ver_combo.currentIndexChanged.connect(self._generate_diff_preview)
        config_layout.addWidget(self.base_ver_combo, stretch=1)

        arrow_lbl = QLabel("➔")
        arrow_lbl.setStyleSheet(f"font-weight: bold; color: {C.TEXT_MUTED}; font-size: 14px;")
        config_layout.addWidget(arrow_lbl)

        config_layout.addWidget(QLabel("<b>Target Release (New):</b>"))
        self.target_ver_combo = QComboBox()
        self.target_ver_combo.currentIndexChanged.connect(self._generate_diff_preview)
        config_layout.addWidget(self.target_ver_combo, stretch=1)

        clause_lbl = QLabel(f"<b>Clause:</b> {self.clause_number} ({self.clause_title})")
        clause_lbl.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        config_layout.addWidget(clause_lbl, stretch=1)

        main_layout.addWidget(config_group)

        # ---------------------------------------------------------------------
        # Main Tabbed Area: [1] Visual Diff  |  [2] AI Analysis Prompt
        # ---------------------------------------------------------------------
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("""
            QTabBar::tab {
                font-weight: bold;
                padding: 6px 16px;
            }}
        """)

        # Tab 1: Visual Diff
        self.tab_diff = QWidget()
        diff_layout = QVBoxLayout(self.tab_diff)
        diff_layout.setContentsMargins(8, 8, 8, 8)
        diff_layout.setSpacing(8)

        diff_top_bar = QHBoxLayout()
        self.diff_stats_lbl = QLabel("<b>Diff Summary:</b> Calculating...")
        self.diff_stats_lbl.setStyleSheet(f"font-size: 12px; color: {C.TEXT_STRONG};")
        diff_top_bar.addWidget(self.diff_stats_lbl, stretch=1)

        self.btn_copy_diff = QPushButton("📋 Copy Diff")
        self.btn_copy_diff.setToolTip("Copy raw unified diff text to clipboard.")
        self.btn_copy_diff.setStyleSheet(f"padding: 4px 12px; font-weight: bold; background-color: {C.BG_CANVAS}; border: 1px solid {C.BORDER_STRONG}; border-radius: 4px;")
        self.btn_copy_diff.clicked.connect(self._copy_diff)
        diff_top_bar.addWidget(self.btn_copy_diff)

        self.btn_view_changes = QPushButton(f"📜 v{self.version_b} Changes / TDocs")
        self.btn_view_changes.setToolTip(f"View all integrated CRs, p-CRs, and TDocs that created v{self.version_b}, and directly open the source contribution documents.")
        self.btn_view_changes.setStyleSheet(f"padding: 4px 12px; font-weight: bold; background-color: {C.SUCCESS_SOFT}; color: {C.SUCCESS_DARK}; border: 1px solid {C.SUCCESS_BORDER_MEDIUM}; border-radius: 4px;")
        self.btn_view_changes.setCursor(Qt.PointingHandCursor)
        self.btn_view_changes.clicked.connect(self._open_version_changes)
        diff_top_bar.addWidget(self.btn_view_changes)

        self.btn_find_tdoc = QPushButton("🎯 Identify Modifying TDoc")
        self.btn_find_tdoc.setToolTip("Identify which specific CR or p-CR introduced or modified this clause across releases.")
        self.btn_find_tdoc.setStyleSheet(f"padding: 4px 12px; font-weight: bold; background-color: {C.SUCCESS_SOFT}; color: {C.SUCCESS_DARK}; border: 1px solid {C.SUCCESS_BORDER_MEDIUM}; border-radius: 4px;")
        self.btn_find_tdoc.setCursor(Qt.PointingHandCursor)
        self.btn_find_tdoc.clicked.connect(self._focus_and_run_attribution)
        diff_top_bar.addWidget(self.btn_find_tdoc)

        diff_layout.addLayout(diff_top_bar)

        self.diff_browser = QTextEdit()
        self.diff_browser.setReadOnly(True)
        self.diff_browser.setStyleSheet(f"""
            QTextEdit {{
                background-color: {C.BG_ALT_LIGHT};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11px;
                line-height: 1.45;
                color: {C.TEXT_PRIMARY};
                padding: 8px;
            }}
        """)
        diff_layout.addWidget(self.diff_browser)
        self.tab_widget.addTab(self.tab_diff, "📊 Visual Diff")

        # Tab 2: AI Analysis Prompt
        self.tab_prompt = QWidget()
        prompt_layout = QVBoxLayout(self.tab_prompt)
        prompt_layout.setContentsMargins(8, 8, 8, 8)
        prompt_layout.setSpacing(8)

        prompt_opts_card = QFrame()
        prompt_opts_card.setStyleSheet(f"background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 4px;")
        opts_vlayout = QVBoxLayout(prompt_opts_card)
        opts_vlayout.setContentsMargins(8, 8, 8, 8)
        opts_vlayout.setSpacing(6)

        row_focus = QHBoxLayout()
        row_focus.addWidget(QLabel("<b>Analysis Focus:</b>"))
        self.focus_combo = QComboBox()
        self.focus_combo.addItem("📋 General Standards & Functional Impact", "standards")
        self.focus_combo.addItem("⚖️ Patent & Prior Art Evaluation", "patent")
        self.focus_combo.addItem("📡 Signalling & Protocol Encoding", "signalling")
        self.focus_combo.currentIndexChanged.connect(self._generate_diff_preview)
        row_focus.addWidget(self.focus_combo, stretch=2)
        opts_vlayout.addLayout(row_focus)

        row_scope = QHBoxLayout()
        row_scope.addWidget(QLabel("<b>Context Scope:</b>"))
        self.tier_group = QButtonGroup(self)

        self.radio_tier1 = QRadioButton("Tier 1: Exact Clause Only")
        self.radio_tier1.setToolTip("Includes only the delta of the target clause.")
        self.tier_group.addButton(self.radio_tier1, 1)
        row_scope.addWidget(self.radio_tier1)

        self.radio_tier2 = QRadioButton("Tier 2: + Parent Procedure Scope (Recommended)")
        self.radio_tier2.setToolTip("Includes hierarchical breadcrumbs, parent intro/preconditions, and the clause delta.")
        self.radio_tier2.setChecked(True)
        self.tier_group.addButton(self.radio_tier2, 2)
        row_scope.addWidget(self.radio_tier2)

        self.radio_tier3 = QRadioButton("Tier 3: + Full Procedure Branch")
        self.radio_tier3.setToolTip("Includes sibling subclauses in the same procedure branch for deep architectural context.")
        self.tier_group.addButton(self.radio_tier3, 3)
        row_scope.addWidget(self.radio_tier3)

        row_scope.addStretch()
        self.tier_group.buttonClicked.connect(self._generate_diff_preview)
        opts_vlayout.addLayout(row_scope)

        prompt_layout.addWidget(prompt_opts_card)

        self.prompt_browser = QTextEdit()
        self.prompt_browser.setReadOnly(True)
        self.prompt_browser.setStyleSheet(f"""
            QTextEdit {{
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11px;
                line-height: 1.4;
                color: {C.TEXT_STRONG};
                padding: 6px;
            }}
        """)
        prompt_layout.addWidget(self.prompt_browser)

        # Backward compatibility alias
        self.preview_browser = self.prompt_browser

        prompt_actions = QHBoxLayout()
        self.btn_copy_prompt = QPushButton("📋 Copy Prompt for LLM")
        self.btn_copy_prompt.setStyleSheet(f"font-weight: bold; background-color: {C.ACCENT_SKY}; color: white; padding: 6px 16px; border-radius: 4px;")
        self.btn_copy_prompt.clicked.connect(self._copy_prompt)
        prompt_actions.addWidget(self.btn_copy_prompt)

        self.btn_save = QPushButton("💾 Save Markdown (.md)")
        self.btn_save.setStyleSheet(f"font-weight: bold; background-color: {C.BG_CANVAS}; color: {C.TEXT_BODY}; border: 1px solid {C.BORDER_STRONG}; padding: 6px 14px; border-radius: 4px;")
        self.btn_save.clicked.connect(self._save_markdown)
        prompt_actions.addWidget(self.btn_save)

        prompt_actions.addStretch()
        prompt_layout.addLayout(prompt_actions)

        self.tab_widget.addTab(self.tab_prompt, "🤖 AI Analysis Prompt")

        # Tab 3: Modifying TDoc Attribution
        self._setup_attribution_tab()
        self.tab_widget.currentChanged.connect(self._on_tab_changed)

        main_layout.addWidget(self.tab_widget)

        # ---------------------------------------------------------------------
        # Bottom Bar: Word Comparison & Close
        # ---------------------------------------------------------------------
        btn_bar = QHBoxLayout()

        self.btn_word_compare = QPushButton("📝 Open in Word (Track Changes)")
        self.btn_word_compare.setToolTip(
            "Extract this clause from Base and Target releases and open Microsoft Word with native Track Changes (strikethrough deletions & underlined additions)."
        )
        self.btn_word_compare.setStyleSheet(f"""
            QPushButton {{
                font-weight: bold;
                background-color: {C.DIFF_ADDED_BG};
                color: {C.SUCCESS};
                border: 1px solid {C.SUCCESS_BORDER_MEDIUM};
                padding: 6px 16px;
                border-radius: 4px;
            }}
            QPushButton:hover {{
                background-color: {C.SUCCESS_SOFT_ALT};
                color: {C.SUCCESS_DARK};
            }}
            QPushButton:disabled {{
                background-color: {C.BG_CANVAS};
                color: {C.TEXT_FAINT};
                border: 1px solid {C.BORDER_SUBTLE};
            }}
        """)
        self.btn_word_compare.clicked.connect(self._start_word_comparison)
        btn_bar.addWidget(self.btn_word_compare)

        btn_bar.addStretch()
        self.btn_close = QPushButton("Close")
        self.btn_close.setStyleSheet("padding: 6px 18px;")
        self.btn_close.clicked.connect(self.accept)
        btn_bar.addWidget(self.btn_close)

        main_layout.addLayout(btn_bar)
