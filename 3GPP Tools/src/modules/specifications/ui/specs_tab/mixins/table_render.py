"""Specifications catalogue table rendering and version badge mixin."""

from __future__ import annotations

import logging
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QMenu,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import scale_px, scale_row_height
from core.ui.theme import C
from core.ui.ui_components import get_wg_badge_style_and_desc
from modules.specifications.ui.components import HoverMenuButton
from modules.specifications.ui.spec_menu import populate_spec_copy_menu
from ..constants import PART_NUM_REGEX

logger = logging.getLogger(__name__)


class SpecsTabTableRenderMixin:
    """Populates filtered specification rows, WG/tag badges, and Column 0 action menus."""

    def _update_combo_item_badges(self, combo: QComboBox, index: int):
        """Update the display text, badges, and tooltips of a combo item based on disk state."""
        if index < 0 or index >= combo.count():
            return
        c_data = combo.itemData(index)
        if not c_data:
            return

        spec_target_dir = Path(self.download_dir) / c_data["spec_num"]
        stem = Path(c_data["fname"]).stem.lower()
        zip_exists = (spec_target_dir / c_data["fname"]).exists()

        local_files = []
        if spec_target_dir.exists():
            try:
                local_files = [f.name.lower() for f in spec_target_dir.iterdir() if f.is_file()]
            except Exception:
                local_files = []

        word_exists = any(f.startswith(stem) and (f.endswith(".docx") or f.endswith(".doc")) for f in local_files)
        pdf_exists = any(f.startswith(stem) and f.endswith(".pdf") for f in local_files)
        html_exists = any(f.startswith(stem) and f.endswith(".html") for f in local_files)
        md_exists = any(f.startswith(stem) and f.endswith(".md") for f in local_files)
        md_split_exists = (spec_target_dir / f"{stem}_clauses" / "index.md").exists()

        c_data["is_downloaded"] = zip_exists
        c_data["has_word"] = word_exists
        c_data["has_pdf"] = pdf_exists
        c_data["has_html"] = html_exists
        c_data["has_md"] = md_exists
        c_data["has_md_split"] = md_split_exists

        badges = []
        tooltips = []
        if word_exists:
            badges.append("📝")
            tooltips.append("Word (.docx) ready")
        if pdf_exists:
            badges.append("📕")
            tooltips.append("PDF (.pdf) ready")
        if html_exists:
            badges.append("🌐")
            tooltips.append("HTML (.html) ready")
        if md_exists:
            badges.append("📘")
            tooltips.append("Markdown (.md) ready")
        if md_split_exists:
            badges.append("📑")
            tooltips.append("Markdown Clauses ready")

        if not badges and zip_exists:
            badges.append("📦")
            tooltips.append("ZIP archive downloaded")
        elif not zip_exists:
            tooltips.append("Available on 3GPP FTP")

        status = "✅ " if zip_exists else ""
        u_date = c_data.get("upload_date")
        date_label = f" ({u_date})" if u_date else ""
        badge_str = f"  {' '.join(badges)}" if badges else ""
        ver = c_data.get("version", "")

        new_text = f"{status}v{ver}{date_label}{badge_str}"
        combo.setItemText(index, new_text)
        combo.setItemData(index, c_data)
        combo.setItemData(index, " • ".join(tooltips), Qt.ToolTipRole)

    def refresh_table(self):
        self.search_timer.stop()
        self.filter_timer.stop()
        self.table.setUpdatesEnabled(False)
        self.table.blockSignals(True)
        try:
            spec_query = self.spec_search_input.text().strip()
            version_query = self.version_search_input.text().strip()
            downloaded_only = self.downloaded_only_cb.isChecked()
            base_dl_dir = Path(self.download_dir)

            is_filtered = any(
                [
                    self._current_tsg,
                    self.table_filters["series"],
                    self.table_filters["tech"],
                    self.table_filters["group"],
                    self.table_filters["spec_type"] != "Any",
                ]
            )

            specs = self.db.search_files(
                spec_number=spec_query if spec_query else None,
                release_version=version_query if version_query else None,
                tag=self.tag_filter,
                tsg=self._current_tsg if self._current_tsg else None,
                **self.table_filters,
            )

            history_summary = {}
            if hasattr(self, "history_db") and self.history_db:
                try:
                    history_summary = self.history_db.get_version_summary()
                except Exception as e:
                    logger.debug(f"Could not prefetch history summary: {e}")

            self.table.setRowCount(0)

            grouped_specs = {}
            for row in specs:
                series, spec_num, title, spec_type, filename, version, url, upload_date, *rest = row
                row_tags = rest[0] if len(rest) > 0 else ""
                row_comments = rest[1] if len(rest) > 1 else ""
                row_wg = rest[2] if len(rest) > 2 else ""
                if filename and "-" not in spec_num:
                    part_match = PART_NUM_REGEX.search(filename)
                    if part_match:
                        spec_num = f"{spec_num}-{part_match.group(1)}"

                if spec_num not in grouped_specs:
                    grouped_specs[spec_num] = {
                        "title": title,
                        "type": spec_type if spec_type else "TS",
                        "wg": row_wg or "",
                        "versions": [],
                        "tags": [t.strip() for t in row_tags.split(",") if t.strip()] if row_tags else [],
                        "comments": row_comments or "",
                    }
                elif not grouped_specs[spec_num].get("wg") and row_wg:
                    grouped_specs[spec_num]["wg"] = row_wg
                grouped_specs[spec_num]["versions"].append((version, url, filename, upload_date))

            if downloaded_only and base_dl_dir.exists():
                try:
                    existing_dirs = {d.name for d in base_dl_dir.iterdir() if d.is_dir()}
                    grouped_specs = {s_num: s_data for s_num, s_data in grouped_specs.items() if s_num in existing_dirs}
                except Exception:
                    pass

            if self.favorites_only:
                filtered_grouped = {}
                for s_num, s_data in grouped_specs.items():
                    if s_num in self.favorites:
                        filtered_grouped[s_num] = s_data
                grouped_specs = filtered_grouped

            total_found = len(grouped_specs)
            rendered_specs = list(grouped_specs.items())[:100]

            if self.favorites_only:
                self.count_badge.setText(f"⭐ {total_found} Favorites")
                self.count_badge.setStyleSheet(f"""
                    padding: 2px 10px; font-size: 11px; font-weight: bold;
                    background-color: {C.WARNING_SOFT_ALT}; color: {C.WARNING_DARK}; border: 1px solid {C.WARNING_ACCENT}; border-radius: 10px;
                """)
            elif not spec_query and not version_query and not is_filtered and not downloaded_only:
                self.count_badge.setText(f"Top {len(rendered_specs)} specifications")
                self.count_badge.setStyleSheet(f"""
                    padding: 2px 10px; font-size: 11px; font-weight: bold;
                    background-color: {C.INFO_SOFT_ALT}; color: {C.INFO_ALT}; border: 1px solid {C.INFO_BORDER_ALT}; border-radius: 10px;
                """)
            elif total_found > 100:
                self.count_badge.setText(f"Showing 100 of {total_found} specs")
                self.count_badge.setStyleSheet(f"""
                    padding: 2px 10px; font-size: 11px; font-weight: bold;
                    background-color: {C.WARNING_SOFT}; color: {C.ORANGE_DARK}; border: 1px solid {C.WARNING_BORDER_STRONG}; border-radius: 10px;
                """)
            else:
                self.count_badge.setText(f"{total_found} specifications")
                self.count_badge.setStyleSheet(f"""
                    padding: 2px 10px; font-size: 11px; font-weight: bold;
                    background-color: {C.BG_CANVAS}; color: {C.TEXT_SECONDARY}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 10px;
                """)

            self._update_active_tags()

            for row_idx, (spec_num, data) in enumerate(rendered_specs):
                self.table.insertRow(row_idx)
                for col_idx in (0, 1, 2):
                    placeholder = QTableWidgetItem("")
                    placeholder.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    self.table.setItem(row_idx, col_idx, placeholder)
                spec_target_dir = base_dl_dir / spec_num
                is_fav = spec_num in self.favorites

                # --- COLUMN 0: Action Button + Badged Spec Number ---
                spec_widget = QWidget()
                spec_widget.setObjectName("specRowCell")
                spec_layout = QHBoxLayout(spec_widget)
                spec_layout.setContentsMargins(4, 2, 4, 2)
                spec_layout.setSpacing(5)
                spec_layout.setAlignment(Qt.AlignVCenter)

                action_btn = HoverMenuButton("⋮")
                action_btn.setFixedSize(scale_px(20), scale_px(22))
                action_btn.setToolTip("Specification Actions")
                action_btn.setCursor(Qt.PointingHandCursor)
                action_btn.setStyleSheet(f"""
                    QPushButton {{ border: none; background: transparent; color: {C.TEXT_MUTED}; font-size: 18px; font-weight: bold; padding-bottom: 3px; }}
                    QPushButton:hover {{ color: {C.ACCENT_PRIMARY}; }}
                    QPushButton::menu-indicator {{ image: none; width: 0px; }}
                """)

                menu = QMenu(self)
                menu.setStyleSheet(f"""
                    QMenu {{ background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; }}
                    QMenu::item {{ padding: 5px 20px 5px 15px; color: {C.TEXT_BODY}; }}
                    QMenu::item:selected {{ background-color: {C.ACCENT_SOFT}; color: {C.ACCENT_PRIMARY}; }}
                    QMenu::item:disabled {{ color: {C.TEXT_FAINT}; }}
                """)

                fav_label = "⭐ Remove from Favorites" if is_fav else "⭐ Add to Favorites"
                menu.addAction(fav_label).triggered.connect(lambda _, s=spec_num: self._toggle_favorite(s))

                ann_action = menu.addAction("🏷️  Edit Tags && Notes...")
                ann_action.triggered.connect(lambda _, s=spec_num: self._open_edit_annotations(s))

                menu.addSeparator()

                info_action = menu.addAction("ℹ️  View Details")
                info_action.triggered.connect(lambda _, s=spec_num: self._show_spec_info(s))

                web_action = menu.addAction("🌐  Open 3GPP Web Report")
                web_action.triggered.connect(lambda _, s=spec_num: self._open_web_report(s))

                search_action = menu.addAction("🔎  Track Evolution in Spec Search")
                search_action.triggered.connect(lambda _, s=spec_num, r=row_idx: self._handle_search_spec_action(s, r))

                index_action = menu.addAction("📥  Index into Spec Search DB")
                index_action.triggered.connect(lambda _, s=spec_num, r=row_idx: self._handle_index_spec_action(s, r))

                menu.addSeparator()

                # --- Quick Copy Submenu ---
                copy_sub = menu.addMenu("📋  Quick Copy")
                copy_sub.setToolTipsVisible(True)
                copy_sub.setToolTip("Copy specification metadata in plain text and rich HTML form.")
                copy_sub.menuAction().setToolTip("Copy specification metadata in plain text and rich HTML form.")

                def _on_action_copy_about_to_show(sub=copy_sub, s=spec_num, r=row_idx):
                    sub.clear()
                    cell_widget = self.table.cellWidget(r, 2)
                    v_info = None
                    if cell_widget:
                        combo = cell_widget.findChild(QComboBox)
                        if combo:
                            v_info = combo.currentData()
                    populate_spec_copy_menu(sub, self, s, version_info=v_info, db=self.db)

                copy_sub.aboutToShow.connect(_on_action_copy_about_to_show)

                menu.addSeparator()

                folder_action = menu.addAction("📂  Open Local Folder")
                folder_action.triggered.connect(lambda _, s=spec_num: self._open_spec_folder(s))

                def _update_menu_state(act=folder_action, path=spec_target_dir, idx_act=index_action, s=spec_num, r=row_idx, csub=copy_sub):
                    if path.exists() and any(path.iterdir()):
                        act.setText("📂  Open Local Folder")
                        act.setEnabled(True)
                    else:
                        act.setText("📁  Folder Not Created")
                        act.setEnabled(False)
                    ver = self._get_row_version(r)
                    if ver:
                        csub.setTitle(f"📋  Quick Copy (v{ver})")
                        is_idx = (
                            self.spec_search_db.is_version_indexed(s, ver)
                            if hasattr(self, "spec_search_db") and self.spec_search_db
                            else False
                        )
                        if is_idx:
                            idx_act.setText(f"🔄  Re-index v{ver} in Spec Search DB")
                        else:
                            idx_act.setText(f"📥  Index v{ver} into Spec Search DB")
                    else:
                        csub.setTitle("📋  Quick Copy")
                        idx_act.setText("📥  Index into Spec Search DB")

                menu.addSeparator()
                del_action = menu.addAction("🗑️  Delete Specification")
                del_action.triggered.connect(lambda _, s=spec_num: self._delete_single_spec(s))

                menu.aboutToShow.connect(_update_menu_state)
                _update_menu_state()
                action_btn.setMenu(menu)

                spec_type = (data["type"] or "TS").upper()
                wg_name = (data.get("wg") or "").strip()
                wg_badge = QLabel(wg_name if wg_name and wg_name != "-" else "—")
                wg_badge.setMinimumWidth(max(48, scale_px(48)))
                wg_badge.setAlignment(Qt.AlignCenter)
                if wg_name and wg_name != "-":
                    wg_badge.setCursor(Qt.PointingHandCursor)
                    wg_style, wg_desc = get_wg_badge_style_and_desc(wg_name)
                    wg_badge.setStyleSheet(wg_style)
                    wg_badge.setToolTip(f"Working Group: {wg_name} ({wg_desc})\nClick to filter by this Working Group")
                    wg_badge.mousePressEvent = lambda event, wg=wg_name: self._on_wg_badge_clicked(wg)
                else:
                    wg_badge.setStyleSheet(f"""
                        background-color: {C.BG_SURFACE}; color: {C.TEXT_FAINT}; border: 1px solid {C.BORDER_SUBTLE};
                        border-radius: 3px; padding: 2px 4px; font-size: 11px; font-weight: bold;
                    """)
                    wg_badge.setToolTip("Working Group: Unassigned / Plenary")

                canonical_label = f"{spec_type} {spec_num}"
                display_title = f"{canonical_label} ⭐" if is_fav else canonical_label
                spec_label = QLabel(display_title)
                spec_label.setObjectName("specNumberLabel")
                # Keep the canonical identifier separate from decorated display text.
                spec_label.setProperty("specNumber", spec_num)
                spec_label.setStyleSheet(f"font-size: 12px; font-weight: bold; color: {C.TEXT_PRIMARY};")
                if is_fav:
                    spec_label.setToolTip("Marked as Favorite")

                spec_layout.addWidget(action_btn)
                spec_layout.addWidget(wg_badge)
                spec_layout.addWidget(spec_label)
                spec_layout.addStretch()
                self.table.setCellWidget(row_idx, 0, spec_widget)

                # --- COLUMN 1: Title + Tags & Comments ---
                title_text = data["title"] if data["title"] else "Unknown Title"
                tags = data.get("tags", [])
                comments = data.get("comments", "").strip()

                if tags or comments:
                    title_widget = QWidget()
                    title_widget.setObjectName("specRowCell")
                    title_layout = QVBoxLayout(title_widget)
                    title_layout.setContentsMargins(6, 6, 6, 6)
                    title_layout.setSpacing(4)

                    title_lbl = QLabel(title_text)
                    title_lbl.setStyleSheet(f"font-size: 12px; color: {C.TEXT_STRONG}; background: transparent;")
                    title_lbl.setWordWrap(True)
                    title_layout.addWidget(title_lbl)

                    meta_row = QHBoxLayout()
                    meta_row.setContentsMargins(0, 0, 0, 0)
                    meta_row.setSpacing(5)

                    for tag_str in tags:
                        tag_pill = QLabel(f"🏷️ {tag_str}")
                        tag_pill.setStyleSheet(f"""
                            background-color: {C.ACCENT_SOFT}; color: {C.ACCENT_PRIMARY}; border: 1px solid {C.ACCENT_RING};
                            border-radius: 9px; padding: 2px 7px; font-size: 11px; font-weight: bold;
                        """)
                        tag_pill.setToolTip(f"User Tag: {tag_str}")
                        meta_row.addWidget(tag_pill)

                    if comments:
                        note_pill = QLabel("📝 Note")
                        note_pill.setStyleSheet(f"""
                            background-color: {C.WARNING_SOFT_ALT}; color: {C.WARNING_DARK}; border: 1px solid {C.WARNING_BORDER_STRONG};
                            border-radius: 9px; padding: 2px 7px; font-size: 11px; font-weight: bold;
                        """)
                        preview = comments if len(comments) <= 250 else f"{comments[:247]}..."
                        note_pill.setToolTip(f"Personal Note:\n{preview}")
                        meta_row.addWidget(note_pill)

                    meta_row.addStretch()
                    title_layout.addLayout(meta_row)
                    self.table.setCellWidget(row_idx, 1, title_widget)
                    # Give generous row height so wrapped titles and 11px tag pills never clip at the bottom edge.
                    needed_h = max(scale_row_height(64), title_widget.sizeHint().height() + scale_px(12))
                    if len(title_text) > 55:
                        needed_h = max(needed_h, scale_row_height(78))
                    self.table.setRowHeight(row_idx, needed_h)
                else:
                    self.table.setItem(row_idx, 1, QTableWidgetItem(title_text))
                    self.table.setRowHeight(row_idx, scale_row_height(44))

                self._build_spec_row_documents_cell(row_idx, spec_num, data, spec_target_dir, history_summary)

        except Exception as e:
            logger.error(f"Error during refresh_table: {e}", exc_info=True)
        finally:
            self.table.blockSignals(False)
            self.table.setUpdatesEnabled(True)
