"""Uniform provider/model/effort menu and defaults dialog for AI email summaries."""

from __future__ import annotations

from typing import Callable, Dict, List, Optional

from PyQt5.QtWidgets import (
    QComboBox,
    QDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.ai.llm_runner import (
    PROVIDER_META,
    SUPPORTED_PROVIDERS,
    ModelSpec,
    ProviderModelEntry,
    discover_all_runner_models,
    get_default_model_spec,
    resolve_effective_model_spec,
    set_default_model_spec,
)
from core.ui.font_scale import scale_px
from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY, escape_qt_mnemonic


def _spec_matches(active: Optional[ModelSpec], prov: str, model: str, effort: str = "") -> bool:
    if active is None:
        return False
    a = active.normalized()
    return a.provider == prov and a.model == model and a.reasoning_effort == effort


def populate_summary_model_menu(
    menu: QMenu,
    *,
    role: str = "leaf",
    active_spec: Optional[ModelSpec] = None,
    on_selected: Optional[Callable[[ModelSpec], None]] = None,
    on_defaults_saved: Optional[Callable[[], None]] = None,
    parent: Optional[QWidget] = None,
    grouped_models: Optional[Dict[str, List[ProviderModelEntry]]] = None,
) -> QMenu:
    """Populate a uniform QMenu listing Ollama, Codex, Claude, and Gemini models + efforts."""
    menu.clear()
    menu.setToolTipsVisible(True)
    eff_active = active_spec or resolve_effective_model_spec(role=role)
    models_by_prov = grouped_models if grouped_models is not None else discover_all_runner_models()

    def _pick(prov: str, slug: str, effort: str = "") -> None:
        chosen = ModelSpec(provider=prov, model=slug, reasoning_effort=effort).normalized()
        if on_selected is not None:
            on_selected(chosen)

    for prov in SUPPORTED_PROVIDERS:
        label, icon = PROVIDER_META[prov]
        entries = models_by_prov.get(prov) or []
        if not entries:
            act_off = menu.addAction(escape_qt_mnemonic(f"{icon} {label} (Offline / Not Installed)"))
            act_off.setEnabled(False)
            continue

        sub = menu.addMenu(escape_qt_mnemonic(f"{icon} {label}"))
        sub.setToolTipsVisible(True)

        if prov != "ollama":
            act_def = sub.addAction("🔄 CLI Default (No Override)")
            act_def.setCheckable(True)
            act_def.setChecked(_spec_matches(eff_active, prov, "", ""))
            act_def.triggered.connect(lambda _=False, p=prov: _pick(p, "", ""))
            sub.addSeparator()

        for ent in entries:
            base_lbl = (
                f"{ent.display_name} ({ent.model_slug})"
                if ent.display_name and ent.display_name != ent.model_slug
                else ent.model_slug
            )
            if ent.reasoning_levels:
                is_mod = eff_active is not None and eff_active.provider == prov and eff_active.model == ent.model_slug
                m_lbl = f"{base_lbl} [{eff_active.reasoning_effort}]" if (is_mod and eff_active.reasoning_effort) else base_lbl
                m_sub = sub.addMenu(escape_qt_mnemonic(m_lbl))
                act_m_def = m_sub.addAction(
                    f"Default Effort ({ent.default_effort})" if ent.default_effort else "Default Effort"
                )
                act_m_def.setCheckable(True)
                act_m_def.setChecked(_spec_matches(eff_active, prov, ent.model_slug, ""))
                act_m_def.triggered.connect(lambda _=False, p=prov, s=ent.model_slug: _pick(p, s, ""))
                m_sub.addSeparator()
                for lvl in ent.reasoning_levels:
                    act_lvl = m_sub.addAction(escape_qt_mnemonic(lvl))
                    act_lvl.setCheckable(True)
                    act_lvl.setChecked(_spec_matches(eff_active, prov, ent.model_slug, lvl))
                    act_lvl.triggered.connect(lambda _=False, p=prov, s=ent.model_slug, lv=lvl: _pick(p, s, lv))
            else:
                act_m = sub.addAction(escape_qt_mnemonic(base_lbl))
                act_m.setCheckable(True)
                act_m.setChecked(_spec_matches(eff_active, prov, ent.model_slug, ""))
                act_m.triggered.connect(lambda _=False, p=prov, s=ent.model_slug: _pick(p, s, ""))

    menu.addSeparator()
    act_cfg = menu.addAction("⚙️ Configure Summary Model Defaults...")
    act_cfg.setToolTip("Set default Fast/Leaf and Deep/Synthesis models across Ollama and external CLIs.")

    def _open_defaults():
        dlg = SummaryModelDefaultsDialog(parent=parent)
        if dlg.exec_() == QDialog.Accepted and on_defaults_saved is not None:
            on_defaults_saved()

    act_cfg.triggered.connect(_open_defaults)
    return menu


class _RoleSelectorWidget(QFrame):
    """Uniform Provider -> Model -> Effort selector row for one summary role."""

    def __init__(
        self,
        title: str,
        subtitle: str,
        role: str,
        models_by_prov: Dict[str, List[ProviderModelEntry]],
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.role = role
        self._models_by_prov = models_by_prov
        self.setObjectName(f"roleCard_{role}")
        self.setStyleSheet(
            f"#roleCard_{role} {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 8, 10, 8)
        lay.setSpacing(6)

        hdr = QLabel(f"<b>{title}</b> <span style='color:{C.TEXT_MUTED}; font-size:11px;'>— {subtitle}</span>")
        hdr.setWordWrap(True)
        lay.addWidget(hdr)

        form = QFormLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setSpacing(6)

        self.prov_combo = QComboBox(self)
        self.prov_combo.setMinimumHeight(scale_px(26))
        self.model_combo = QComboBox(self)
        self.model_combo.setMinimumHeight(scale_px(26))
        self.effort_combo = QComboBox(self)
        self.effort_combo.setMinimumHeight(scale_px(26))

        form.addRow("Provider:", self.prov_combo)
        form.addRow("Model:", self.model_combo)
        form.addRow("Effort:", self.effort_combo)
        lay.addLayout(form)

        self.prov_combo.currentIndexChanged.connect(self._on_provider_changed)
        self.model_combo.currentIndexChanged.connect(self._on_model_changed)
        self.load_from_config(models_by_prov)

    def load_from_config(self, models_by_prov: Dict[str, List[ProviderModelEntry]]) -> None:
        self._models_by_prov = models_by_prov
        saved = get_default_model_spec(self.role)
        self.prov_combo.blockSignals(True)
        self.prov_combo.clear()
        self.prov_combo.addItem("🔄 Auto (Smart Fallback)", "")
        for prov in SUPPORTED_PROVIDERS:
            label, icon = PROVIDER_META[prov]
            entries = models_by_prov.get(prov) or []
            suffix = f" ({len(entries)} models)" if entries else " (offline / not found)"
            self.prov_combo.addItem(f"{icon} {label}{suffix}", prov)
        if saved and saved.provider:
            idx = self.prov_combo.findData(saved.provider)
            if idx >= 0:
                self.prov_combo.setCurrentIndex(idx)
        self.prov_combo.blockSignals(False)
        self._populate_models(selected_model=saved.model if saved else "", selected_effort=saved.reasoning_effort if saved else "")

    def _on_provider_changed(self) -> None:
        self._populate_models(selected_model="", selected_effort="")

    def _populate_models(self, *, selected_model: str, selected_effort: str) -> None:
        prov = str(self.prov_combo.currentData() or "").strip()
        self.model_combo.blockSignals(True)
        self.model_combo.clear()
        if not prov:
            self.model_combo.addItem("(Auto-selected)", "")
            self.model_combo.setEnabled(False)
        else:
            self.model_combo.setEnabled(True)
            if prov != "ollama":
                self.model_combo.addItem("(CLI Default)", "")
            for ent in self._models_by_prov.get(prov) or []:
                lbl = f"{ent.display_name} ({ent.model_slug})" if ent.display_name != ent.model_slug else ent.model_slug
                self.model_combo.addItem(lbl, ent.model_slug)
            if selected_model:
                idx = self.model_combo.findData(selected_model)
                if idx >= 0:
                    self.model_combo.setCurrentIndex(idx)
                else:
                    self.model_combo.addItem(selected_model, selected_model)
                    self.model_combo.setCurrentIndex(self.model_combo.count() - 1)
        self.model_combo.blockSignals(False)
        self._populate_efforts(selected_effort=selected_effort)

    def _on_model_changed(self) -> None:
        cur_eff = str(self.effort_combo.currentData() or "").strip()
        self._populate_efforts(selected_effort=cur_eff)

    def _populate_efforts(self, *, selected_effort: str) -> None:
        prov = str(self.prov_combo.currentData() or "").strip()
        slug = str(self.model_combo.currentData() or "").strip()
        levels = ()
        for ent in self._models_by_prov.get(prov) or []:
            if ent.model_slug == slug:
                levels = ent.reasoning_levels
                break
        if not levels and prov in ("codex", "claude"):
            for ent in self._models_by_prov.get(prov) or []:
                if ent.reasoning_levels:
                    levels = ent.reasoning_levels
                    break
        self.effort_combo.clear()
        self.effort_combo.addItem("(Default)", "")
        for lvl in levels:
            self.effort_combo.addItem(lvl, lvl)
        if selected_effort:
            idx = self.effort_combo.findData(selected_effort)
            if idx >= 0:
                self.effort_combo.setCurrentIndex(idx)
            else:
                self.effort_combo.addItem(selected_effort, selected_effort)
                self.effort_combo.setCurrentIndex(self.effort_combo.count() - 1)
        self.effort_combo.setEnabled(bool(prov) and self.effort_combo.count() > 1)

    def get_model_spec(self) -> Optional[ModelSpec]:
        prov = str(self.prov_combo.currentData() or "").strip().lower()
        if not prov:
            return None
        return ModelSpec(
            provider=prov,
            model=str(self.model_combo.currentData() or "").strip(),
            reasoning_effort=str(self.effort_combo.currentData() or "").strip().lower(),
        ).normalized()


class SummaryModelDefaultsDialog(QDialog):
    """Modal configuration dialog for default Fast/Leaf and Deep/Synthesis summary models."""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("AI Email Summary Model Defaults")
        self.resize(560, 420)
        self.setModal(True)

        models_by_prov = discover_all_runner_models(force_refresh=False)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        intro = QLabel(
            "Choose default engines and models for email summarization. "
            "Ollama and external CLIs (Codex, Claude, Gemini) run through the same unified pipeline."
        )
        intro.setWordWrap(True)
        intro.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px;")
        layout.addWidget(intro)

        self.leaf_card = _RoleSelectorWidget(
            "⚡ Fast / Leaf Summaries",
            "Used for Per-Thread, Per-TDoc, and incremental updates",
            "leaf",
            models_by_prov,
            self,
        )
        layout.addWidget(self.leaf_card)

        self.synth_card = _RoleSelectorWidget(
            "🧠 Deep / Synthesis Briefings",
            "Used for Per-AI cross-thinking (Stage 2) and multi-row selections",
            "synthesis",
            models_by_prov,
            self,
        )
        layout.addWidget(self.synth_card)

        btn_row = QHBoxLayout()
        self.rescan_btn = QPushButton("🔄 Re-scan Providers", self)
        self.rescan_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.rescan_btn.clicked.connect(self._on_rescan)
        btn_row.addWidget(self.rescan_btn)
        btn_row.addStretch()

        cancel_btn = QPushButton("Cancel", self)
        cancel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)

        save_btn = QPushButton("💾 Save Defaults", self)
        save_btn.setStyleSheet(
            f"QPushButton {{ background-color: {C.ACCENT_PRIMARY}; color: {C.TEXT_ON_ACCENT}; "
            f"border: none; border-radius: 4px; padding: 6px 14px; font-weight: bold; }} "
            f"QPushButton:hover {{ background-color: {C.ACCENT_HOVER}; }}"
        )
        save_btn.clicked.connect(self._on_save)
        btn_row.addWidget(save_btn)
        layout.addLayout(btn_row)

    def _on_rescan(self) -> None:
        fresh = discover_all_runner_models(force_refresh=True)
        self.leaf_card.load_from_config(fresh)
        self.synth_card.load_from_config(fresh)

    def _on_save(self) -> None:
        set_default_model_spec("leaf", self.leaf_card.get_model_spec())
        set_default_model_spec("synthesis", self.synth_card.get_model_spec())
        self.accept()
