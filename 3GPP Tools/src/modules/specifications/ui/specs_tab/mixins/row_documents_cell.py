"""Column 2 version dropdown and document action menu builder mixin for SpecificationsTab."""

from __future__ import annotations

from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QComboBox, QHBoxLayout, QMenu, QPushButton, QWidget

from core.ui.font_scale import scale_px
from core.ui.theme import C
from modules.specifications.ui.spec_menu import populate_spec_copy_menu


class SpecsTabRowDocumentsCellMixin:
    """Constructs the Column 2 version selector and document action dropdown cell."""

    def _build_spec_row_documents_cell(self, row_idx, spec_num, data, spec_target_dir, history_summary):
        # --- COLUMN 2: Documents Action Bar ---
        version_combo = QComboBox()
        combo_w = max(255, scale_px(255))
        version_combo.setFixedWidth(combo_w)
        version_combo.setFixedHeight(max(30, scale_px(30)))
        version_combo.setStyleSheet(self._version_combo_style)
        version_combo.view().setMinimumWidth(combo_w + 15)

        # Snapshot local files in memory once per specification
        local_files = []
        if spec_target_dir.exists():
            try:
                local_files = [f.name.lower() for f in spec_target_dir.iterdir() if f.is_file()]
            except Exception:
                local_files = []

        def parse_ver(v_str):
            return [(0, int(x)) if x.isdigit() else (1, str(x)) for x in str(v_str).split(".")]

        sorted_versions = sorted(data["versions"], key=lambda x: parse_ver(x[0]), reverse=True)

        for ver, url, fname, u_date in sorted_versions:
            if not fname:
                continue

            stem = Path(fname).stem.lower()
            fname_lower = fname.lower()
            zip_exists = fname_lower in local_files
            word_exists = any(
                f.startswith(stem) and (f.endswith(".docx") or f.endswith(".doc")) for f in local_files
            )
            pdf_exists = any(f.startswith(stem) and f.endswith(".pdf") for f in local_files)
            html_exists = any(f.startswith(stem) and f.endswith(".html") for f in local_files)

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

            hist_cnt = history_summary.get((spec_num, ver), 0)
            if hist_cnt > 0:
                badges.append("📜")
                tooltips.append(f"{hist_cnt} Change History entries (CRs/p-CRs)")

            if not badges and zip_exists:
                badges.append("📦")
                tooltips.append("ZIP archive downloaded")
            elif not zip_exists:
                tooltips.append("Available on 3GPP FTP")

            status = "✅ " if zip_exists else ""
            date_label = f" ({u_date})" if u_date else ""
            badge_str = f"  {' '.join(badges)}" if badges else ""
            display_text = f"{status}v{ver}{date_label}{badge_str}"

            item_idx = version_combo.count()
            version_combo.addItem(
                display_text,
                userData={
                    "url": url,
                    "fname": fname,
                    "spec_num": spec_num,
                    "version": ver,
                    "is_downloaded": zip_exists,
                    "upload_date": u_date,
                    "has_word": word_exists,
                    "has_pdf": pdf_exists,
                    "has_html": html_exists,
                },
            )
            version_combo.setItemData(item_idx, " • ".join(tooltips), Qt.ToolTipRole)

        doc_action_btn = QPushButton()
        doc_action_btn.setFixedWidth(max(150, scale_px(150)))
        doc_action_btn.setFixedHeight(max(30, scale_px(30)))
        doc_action_btn.setCursor(Qt.PointingHandCursor)

        doc_menu = QMenu(self)
        doc_menu.setStyleSheet(f"""
            QMenu {{ background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; }}
            QMenu::item {{ padding: 6px 22px 6px 14px; color: {C.TEXT_BODY}; font-size: 12px; }}
            QMenu::item:selected {{ background-color: {C.ACCENT_SOFT}; color: {C.ACCENT_PRIMARY}; }}
            QMenu::item:disabled {{ color: {C.TEXT_FAINT}; }}
            QMenu::separator {{ height: 1px; background-color: {C.BORDER_SUBTLE}; margin: 4px 0; }}
        """)
        doc_action_btn.setMenu(doc_menu)

        def _update_btn_state(index_ignore=0, c=version_combo, btn=doc_action_btn, menu=doc_menu,
                              files_cache=local_files):
            c_data = c.currentData()
            if not c_data:
                return

            current_dir = Path(self.download_dir) / c_data["spec_num"]
            stem = Path(c_data["fname"]).stem.lower()

            # Re-scan the directory whenever button state is refreshed.
            # The old code used the refresh_table() snapshot (files_cache),
            # which became stale immediately after extracting/converting a file.
            current_files = []
            if current_dir.exists():
                try:
                    current_files = [
                        f.name.lower()
                        for f in current_dir.iterdir()
                        if f.is_file()
                    ]
                except OSError:
                    current_files = []

            zip_exists = c_data["fname"].lower() in current_files
            word_exists = any(
                f.startswith(stem) and (f.endswith(".docx") or f.endswith(".doc"))
                for f in current_files
            )
            pdf_exists = any(f.startswith(stem) and f.endswith(".pdf") for f in current_files)
            html_exists = any(f.startswith(stem) and f.endswith(".html") for f in current_files)
            txt_exists = any(f.startswith(stem) and f.endswith(".txt") for f in current_files)
            md_exists = any(f.startswith(stem) and f.endswith(".md") for f in current_files)
            md_split_exists = (spec_target_dir / f"{stem}_clauses" / "index.md").exists()
            dir_ready = bool(current_files)

            base_style = """
                QPushButton {
                    font-size: 12px;
                    font-weight: bold;
                    padding: 2px 8px;
                    border-radius: 4px;
                    text-align: center;
                }
                QPushButton::menu-indicator { image: none; width: 0px; }
            """

            if word_exists:
                btn.setText("📝 Open Word ▾")
                btn.setStyleSheet(base_style + f"""
                    QPushButton {{
                        background-color: {C.SUCCESS_SOFT};
                        color: {C.SUCCESS};
                        border: 1px solid {C.SUCCESS};
                    }}
                    QPushButton:hover {{ background-color: {C.SUCCESS_SOFT_ALT}; }}
                """)
            elif zip_exists:
                btn.setText("⚙️ Extract Word ▾")
                btn.setStyleSheet(base_style + f"""
                    QPushButton {{
                        background-color: {C.WARNING_SOFT};
                        color: {C.WARNING};
                        border: 1px solid {C.WARNING_BORDER_STRONG};
                    }}
                    QPushButton:hover {{ background-color: {C.ORANGE_SOFT}; }}
                """)
            else:
                btn.setText("⬇️ Get Word ▾")
                btn.setStyleSheet(base_style + f"""
                    QPushButton {{
                        background-color: {C.ACCENT_SOFT};
                        color: {C.ACCENT_BRIGHT};
                        border: 1px solid {C.ACCENT_RING};
                    }}
                    QPushButton:hover {{ background-color: {C.ACCENT_SOFT_HOVER}; border-color: {C.ACCENT_BRIGHT}; }}
                """)

            menu.clear()

            word_label = "📝 Open Word Document" if word_exists else (
                "⚙️ Extract Word Document" if zip_exists else "⬇️ Download && Open Word")
            act_word = menu.addAction(f"{word_label} {'✅' if word_exists else ''}".strip())
            act_word.triggered.connect(lambda: self._handle_document_action(c, "word", btn))

            pdf_label = "📕 Open PDF" if pdf_exists else (
                "⚙️ Convert to PDF" if (word_exists or zip_exists) else "⬇️ Get && Convert to PDF")
            act_pdf = menu.addAction(f"{pdf_label} {'✅' if pdf_exists else ''}".strip())
            act_pdf.triggered.connect(lambda: self._handle_document_action(c, "pdf", btn))

            html_label = "🌐 Open HTML" if html_exists else (
                "⚙️ Convert to HTML" if (word_exists or zip_exists) else "⬇️ Get && Convert to HTML")
            act_html = menu.addAction(f"{html_label} {'✅' if html_exists else ''}".strip())
            act_html.triggered.connect(lambda: self._handle_document_action(c, "html", btn))

            txt_label = "📄 Open TXT" if txt_exists else (
                "⚙️ Convert to TXT" if (word_exists or zip_exists) else "⬇️ Get && Convert to TXT")
            act_txt = menu.addAction(f"{txt_label} {'✅' if txt_exists else ''}".strip())
            act_txt.triggered.connect(lambda: self._handle_document_action(c, "txt", btn))

            menu.addSeparator()

            md_label = "📘 Open Markdown (.md)" if md_exists else (
                "⚙️ Convert to Markdown (.md)" if (word_exists or zip_exists) else "⬇️ Get && Convert to Markdown (.md)")
            act_md = menu.addAction(f"{md_label} {'✅' if md_exists else ''}".strip())
            act_md.triggered.connect(lambda: self._handle_document_action(c, "md", btn))

            md_split_label = "📑 Open Markdown Clauses" if md_split_exists else (
                "⚙️ Convert to Markdown Clauses" if (word_exists or zip_exists) else "⬇️ Get && Convert to Markdown Clauses")
            act_md_split = menu.addAction(f"{md_split_label} {'✅' if md_split_exists else ''}".strip())
            act_md_split.triggered.connect(lambda: self._handle_document_action(c, "md_split", btn))

            menu.addSeparator()

            # Add the new Clause Excerpt action here:
            act_excerpt = menu.addAction("✂️ Extract Clauses (Excerpt)...")
            act_excerpt.triggered.connect(lambda: self._open_clause_excerpt_dialog(c))

            act_cr_pcr = menu.addAction("📝 Draft Contribution (CR / pCR)...")
            act_cr_pcr.triggered.connect(lambda: self._open_cr_pcr_dialog(c))

            menu.addSeparator()

            ver_val = c_data.get("version", "")
            ver_tag = f" (v{ver_val})" if ver_val else ""

            act_history = menu.addAction(f"📜 View Change History && TDocs{ver_tag}...")
            act_history.triggered.connect(lambda _, s=c_data["spec_num"], v=c_data.get("version", ""): self._open_version_history_dialog(s, v))

            act_search_tab = menu.addAction("🔎 Track Evolution / Clauses in Spec Search")
            act_search_tab.triggered.connect(lambda: self.search_in_spec_search_requested.emit(c_data["spec_num"], c_data.get("version", "")))

            is_idx = (
                self.spec_search_db.is_version_indexed(c_data["spec_num"], ver_val)
                if hasattr(self, "spec_search_db") and self.spec_search_db and ver_val
                else False
            )
            idx_lbl = f"🔄 Re-index{ver_tag} in Spec Search DB" if is_idx else f"📥 Index{ver_tag} into Spec Search DB"
            act_index_search = menu.addAction(idx_lbl)
            act_index_search.triggered.connect(lambda: self.index_in_spec_search_requested.emit(c_data["spec_num"], c_data.get("version", "")))

            act_smart_search = menu.addAction("⭐ Index Latest Releases in Spec Search...")
            act_smart_search.setToolTip("Check and ingest the newest version of each release for this specification.")
            act_smart_search.triggered.connect(lambda _, s=c_data["spec_num"]: self.smart_index_in_spec_search_requested.emit(s))

            menu.addSeparator()

            zip_label = "📦 Show ZIP Archive ✅" if zip_exists else "📥 Download ZIP Archive"
            act_zip = menu.addAction(zip_label)
            act_zip.triggered.connect(lambda: self._handle_zip_action(c, btn))

            menu.addSeparator()
            ver_title = f"v{c_data['version']}" if not str(c_data.get('version', '')).startswith('v') else str(c_data.get('version', ''))
            doc_copy_sub = menu.addMenu(f"📋 Quick Copy ({ver_title})")
            doc_copy_sub.setToolTipsVisible(True)
            doc_copy_sub.setToolTip(f"Copy specification metadata and download links for version {ver_title}.")

            def _on_doc_copy_about_to_show(sub=doc_copy_sub, s=c_data["spec_num"], v=c_data):
                sub.clear()
                populate_spec_copy_menu(sub, self, s, version_info=v, db=self.db)

            doc_copy_sub.aboutToShow.connect(_on_doc_copy_about_to_show)

            if dir_ready:
                act_dir = menu.addAction("📂 Open Local Folder")
                act_dir.setEnabled(True)
                act_dir.triggered.connect(lambda _, s=c_data["spec_num"]: self._open_spec_folder(s))
            else:
                act_dir = menu.addAction("📁 Folder Not Created")
                act_dir.setEnabled(False)

        version_combo.currentIndexChanged.connect(_update_btn_state)
        _update_btn_state()

        cell_widget = QWidget()
        cell_widget.setObjectName("specRowCell")
        layout = QHBoxLayout(cell_widget)
        layout.setContentsMargins(6, 4, 6, 4)
        layout.setSpacing(6)
        layout.setAlignment(Qt.AlignVCenter)

        layout.addWidget(version_combo)
        layout.addWidget(doc_action_btn)
        layout.addStretch()

        self.table.setCellWidget(row_idx, 2, cell_widget)
