# --- File: src/modules/meetings/ui/tdocs/window/mixins/table_layout.py ---
"""
TDoc table construction, row heights, fitted/scaled column widths, header sorting, copy selection.

Mixin of TDocsWindow: relies on attributes set in TDocsWindow.__init__ / _setup_*.
"""
from PyQt5.QtCore import (
    Qt,
    QRect,
    QTimer,
)
from PyQt5.QtGui import (
    QCursor,
    QKeySequence,
    QTextDocument,
)
from PyQt5.QtWidgets import (
    QWidget,
    QHBoxLayout,
    QTableView,
    QHeaderView,
    QLabel,
    QApplication,
    QShortcut,
    QToolTip,
)
from core.ui.font_scale import get_font_scale_notifier, scale_px
from modules.meetings.core.settings import MeetingsSettings
from modules.meetings.ui.tdoc_delegates import (
    HtmlDelegate,
    TDocActionDelegate,
    AgendaItemDelegate,
    SourceCompanyDelegate,
)
from modules.meetings.ui.tdocs_components import ComparisonCartWidget
from modules.meetings.ui.tdocs_models import TDocsTableModel, TDocsFilterProxyModel
from core.ui.theme import C


class TableLayoutMixin:
    """TDoc table construction, row heights, fitted/scaled column widths, header sorting, copy selection."""

    def _setup_table(self, layout, data, user_data):
        self.settings = MeetingsSettings()
        self.table = QTableView()
        self.model = TDocsTableModel(self.meeting_dir, data, user_data, agenda_map=self.agenda_map)
        self.proxy = TDocsFilterProxyModel()
        self.proxy.setIncludeUnassignedLsIn(self.settings.get_include_unassigned_ls_in())
        self.proxy.setIncludeTentativeSources(self.settings.get_include_tentative_sources())
        self.proxy.setSourceModel(self.model)
        self.proxy.layoutChanged.connect(self._update_count_label)
        self.proxy.layoutChanged.connect(self._schedule_row_height_adjustment)
        self.proxy.modelReset.connect(self._schedule_row_height_adjustment)
        self.proxy.rowsInserted.connect(lambda *_: self._schedule_row_height_adjustment())
        self.proxy.rowsRemoved.connect(lambda *_: self._schedule_row_height_adjustment())
        self.model.dataChanged.connect(lambda *_: self._schedule_row_height_adjustment())

        self.table.setModel(self.proxy)
        self.table.setWordWrap(True)
        self.table.setSelectionBehavior(QTableView.SelectItems)
        self.table.setSelectionMode(QTableView.ExtendedSelection)
        self.table.doubleClicked.connect(self._show_cell_popup)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setStyleSheet(f"""
            QTableView {{
                gridline-color: {C.BG_CANVAS};
                border: 1px solid {C.BORDER_STRONG};
                background-color: {C.BG_CARD};
            }}
            QHeaderView::section {{
                background-color: {C.BG_SURFACE};
                padding: 4px;
                font-weight: bold;
                border: 1px solid {C.BORDER_SUBTLE};
                color: {C.TEXT_STRONG};
            }}
        """)
        vheader = self.table.verticalHeader()
        vheader.setSectionResizeMode(QHeaderView.Fixed)
        vheader.setMinimumSectionSize(scale_px(48))
        vheader.setDefaultSectionSize(scale_px(48))

        self.action_delegate = TDocActionDelegate(self.table)
        self.action_delegate.actionClicked.connect(self._handle_tdoc_action)
        self.table.setItemDelegateForColumn(0, self.action_delegate)

        self.source_delegate = SourceCompanyDelegate(self.table)
        self.table.setItemDelegateForColumn(3, self.source_delegate)

        self.html_delegate = HtmlDelegate(self.table)
        self.html_delegate.linkClicked.connect(self._scroll_to_tdoc)
        self.html_delegate.linkRightClicked.connect(self._show_related_menu)
        self.table.setItemDelegateForColumn(7, self.html_delegate)
        self.table.setItemDelegateForColumn(12, self.html_delegate)

        self.agenda_delegate = AgendaItemDelegate(self.table)
        self.table.setItemDelegateForColumn(10, self.agenda_delegate)

        header = self.table.horizontalHeader()
        header.setMinimumSectionSize(scale_px(26))
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setSectionResizeMode(6, QHeaderView.Fixed)
        header.setSectionResizeMode(9, QHeaderView.Fixed)
        self._apply_scaled_table_columns()
        header.sectionClicked.connect(self._on_header_section_clicked)
        header.sectionResized.connect(self._on_header_section_resized)
        get_font_scale_notifier().scale_changed.connect(self._on_font_scale_changed)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._on_table_context_menu)

        if self.is_sa2:
            self.order_combo.blockSignals(True)
            self.order_combo.setCurrentIndex(0)  # 📋 Agenda Order
            self.order_combo.blockSignals(False)
            self.proxy.setSortMode("AGENDA")
            self.table.sortByColumn(10, Qt.AscendingOrder)

        layout.addWidget(self.table)

        # Bottom Status & Shared Comparison Cart Bar
        bottom_bar = QHBoxLayout()
        bottom_bar.setContentsMargins(2, 2, 2, 0)
        bottom_bar.setSpacing(8)

        self._footer_status_widget = QWidget(self)
        self._footer_status_widget.setStyleSheet(".QWidget { background: transparent; border: none; }")
        footer_status_layout = QHBoxLayout(self._footer_status_widget)
        footer_status_layout.setContentsMargins(0, 0, 0, 0)
        footer_status_layout.setSpacing(8)
        footer_status_layout.addWidget(self.count_lbl)
        sep_dot = QLabel("·")
        sep_dot.setStyleSheet(f"color: {C.TEXT_FAINT}; font-weight: bold;")
        footer_status_layout.addWidget(sep_dot)
        footer_status_layout.addWidget(self.last_mod_lbl)
        bottom_bar.addWidget(self._footer_status_widget)

        bottom_bar.addStretch()

        self.cart_frame = ComparisonCartWidget(self)
        self.btn_cart_launcher = self.cart_frame.create_launcher_button(self)
        bottom_bar.addWidget(self.btn_cart_launcher)
        bottom_bar.addWidget(self.cart_frame)
        layout.addLayout(bottom_bar)

        self._load_tracked_wi_tokens()
        self._sync_tracked_threads_to_table()
        self._refresh_comboboxes()
        self._refresh_email_counts()
        self._schedule_row_height_adjustment()

        from PyQt5.QtWidgets import QShortcut
        from PyQt5.QtGui import QKeySequence
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.table)
        self.copy_shortcut.activated.connect(self._copy_table_selection)

    def _schedule_row_height_adjustment(self):
        """Coalesce table row-height recalculations onto a short single-shot timer."""
        if hasattr(self, "_row_height_timer") and self._row_height_timer is not None:
            self._row_height_timer.start()

    def _adjust_table_row_heights(self):
        """Expand visible TDocs rows dynamically for wrapped multi-line titles, sources, and remarks."""
        if not hasattr(self, "table") or not hasattr(self, "proxy"):
            return
        row_count = self.proxy.rowCount()
        min_h = scale_px(48)
        max_h = scale_px(124)

        vheader = self.table.verticalHeader()
        vheader.setMinimumSectionSize(min_h)
        vheader.setDefaultSectionSize(min_h)

        if row_count <= 0:
            return

        hheader = self.table.horizontalHeader()
        w_title = max(scale_px(80), hheader.sectionSize(2))
        w_source = max(scale_px(60), hheader.sectionSize(3))
        w_remarks = max(scale_px(140), hheader.sectionSize(7))
        w_related = max(scale_px(80), hheader.sectionSize(12))

        fm = self.table.fontMetrics()
        pad_v = scale_px(8)
        doc = QTextDocument()
        doc.setDocumentMargin(4)
        doc.setDefaultFont(self.table.font())

        has_emails = len(self.model._headers) > 13
        title_wrap_w = max(40, w_title - 10)
        source_wrap_w = max(40, w_source - 10)
        proxy = self.proxy

        for r in range(row_count):
            needed = min_h

            title_txt = str(proxy.data(proxy.index(r, 2), Qt.DisplayRole) or "")
            if title_txt and (len(title_txt) * fm.averageCharWidth() > title_wrap_w or "\n" in title_txt):
                h_t = fm.boundingRect(QRect(0, 0, title_wrap_w, 1000), Qt.TextWordWrap, title_txt).height() + pad_v
                if h_t > needed:
                    needed = h_t

            source_txt = str(proxy.data(proxy.index(r, 3), Qt.DisplayRole) or "")
            if source_txt and (len(source_txt) * fm.averageCharWidth() > source_wrap_w or "\n" in source_txt):
                h_s = fm.boundingRect(QRect(0, 0, source_wrap_w, 1000), Qt.TextWordWrap, source_txt).height() + pad_v
                if h_s > needed:
                    needed = h_s

            remarks_html = str(proxy.data(proxy.index(r, 7), Qt.DisplayRole) or "")
            if remarks_html and ("<" in remarks_html or len(remarks_html) > 35):
                doc.setTextWidth(w_remarks)
                doc.setHtml(remarks_html)
                h_r = int(doc.size().height()) + scale_px(6)
                if h_r > needed:
                    needed = h_r

            if has_emails:
                rel_html = str(proxy.data(proxy.index(r, 12), Qt.DisplayRole) or "")
                if rel_html and len(rel_html) > 25:
                    doc.setTextWidth(w_related)
                    doc.setHtml(rel_html)
                    h_rel = int(doc.size().height()) + scale_px(6)
                    if h_rel > needed:
                        needed = h_rel

            clamped = max(min_h, min(max_h, needed))
            if self.table.rowHeight(r) != clamped:
                self.table.setRowHeight(r, clamped)

    def _on_header_section_resized(self, _logical_index: int, _old_size: int, _new_size: int):
        """Track user-initiated column resizing so responsive relayout preserves manual widths."""
        if not getattr(self, "_programmatic_column_resize", False):
            self._user_resized_cols = True

    @staticmethod
    def compute_fitted_column_widths(viewport_width: int, col_count: int, scale_fn=None) -> dict:
        """
        Computes responsive TDocs table column widths that expand on wide monitors
        and proportionally shrink flexible columns on narrower windows so the
        rightmost Emails unread/flag column remains visible without horizontal clipping.
        """
        sc = scale_fn or scale_px
        base_widths = {
            0: sc(92),
            1: sc(86),
            2: sc(200),
            3: sc(112),
            4: sc(64),
            5: sc(70),
            6: sc(26),
            7: sc(220),
            8: sc(70),
            9: sc(26),
            10: sc(114),
            11: sc(86),
            12: sc(110),
        }
        if col_count > 13:
            base_widths[13] = sc(100)

        min_widths = {
            0: sc(74),
            1: sc(74),
            2: sc(100),
            3: sc(64),
            4: sc(52),
            5: sc(48),
            6: sc(26),
            7: sc(108),
            8: sc(54),
            9: sc(26),
            10: sc(68),
            11: sc(58),
            12: sc(58),
            13: sc(92),
        }

        widths = {col: w for col, w in base_widths.items() if col < col_count}
        total_base = sum(widths.values())
        if viewport_width > total_base > 0:
            surplus = viewport_width - total_base
            rem_extra = int(surplus * 0.52)
            title_extra = int(surplus * 0.34)
            source_extra = surplus - rem_extra - title_extra
            if 7 in widths:
                widths[7] += rem_extra
            if 2 in widths:
                widths[2] += title_extra
            if 3 in widths:
                widths[3] += source_extra
        elif 0 < viewport_width < total_base:
            deficit = total_base - viewport_width
            shrinkable = {
                c: max(0, widths[c] - min_widths.get(c, widths[c]))
                for c in widths
            }
            total_shrinkable = sum(shrinkable.values())
            if total_shrinkable > 0:
                target_cut = min(deficit, total_shrinkable)
                removed = 0
                for c, avail in shrinkable.items():
                    if avail > 0:
                        cut = int(target_cut * avail / total_shrinkable)
                        widths[c] -= cut
                        removed += cut
                # Distribute remaining rounding pixels from widest flexible columns
                leftover = target_cut - removed
                for c in sorted(shrinkable, key=lambda k: widths[k] - min_widths.get(k, widths[k]), reverse=True):
                    if leftover <= 0:
                        break
                    room = widths[c] - min_widths.get(c, widths[c])
                    if room > 0:
                        take = min(room, leftover)
                        widths[c] -= take
                        leftover -= take
        return widths

    def _schedule_column_fit(self):
        """Fits columns on the next event-loop tick so viewport().width() reflects post-layout & scrollbar geometry."""
        if getattr(self, "_user_resized_cols", False):
            return
        if not hasattr(self, "_col_fit_timer"):
            self._col_fit_timer = QTimer(self)
            self._col_fit_timer.setSingleShot(True)
            self._col_fit_timer.timeout.connect(self._apply_scaled_table_columns)
        self._col_fit_timer.start(15)

    def _apply_scaled_table_columns(self):
        """Scale and fit TDocsWindow table column widths to the active viewport."""
        if not hasattr(self, "table") or not hasattr(self, "model"):
            return
        header = self.table.horizontalHeader()
        col_count = len(self.model._headers)
        vp_w = self.table.viewport().width()
        if vp_w > 40:
            # Reserve 2px safety margin so Qt never triggers an unnecessary 1-2px horizontal scrollbar
            vp_w -= 2
            if not self.table.verticalScrollBar().isVisible() and self.proxy.rowCount() > 6:
                vp_w = max(200, vp_w - max(16, self.table.verticalScrollBar().sizeHint().width()))
        widths = self.compute_fitted_column_widths(vp_w, col_count)

        self._programmatic_column_resize = True
        try:
            header.setMinimumSectionSize(scale_px(26))
            for col, width in widths.items():
                if header.sectionSize(col) != width:
                    header.resizeSection(col, width)
        finally:
            self._programmatic_column_resize = False
        self._schedule_row_height_adjustment()

    def _on_font_scale_changed(self, _scale: float = 0.0):
        """Update filter dropdown caps, toolbar buttons, column widths, and row heights on font scale changes."""
        self.instant_fetch_input.setMinimumWidth(scale_px(120))
        self.instant_fetch_input.setMaximumWidth(scale_px(160))
        self.stats_cfg_btn.setMinimumWidth(scale_px(34))
        self.search_input.setMinimumWidth(scale_px(180))
        self.search_input.setMaximumWidth(scale_px(240))
        for btn in (
            self.btn_followed_wis,
            self.btn_starred_wis,
            self.btn_agreements_filter,
            self.btn_unread_emails,
            self.btn_flagged_emails,
        ):
            btn.setFixedSize(scale_px(28), scale_px(28))
        caps = (
            (self.company_combo, 165),
            (self.type_combo, 125),
            (self.ai_combo, 130),
            (self.sub_item_combo, 150),
            (self.wi_combo, 135),
            (self.status_combo, 155),
            (self.my_status_combo, 145),
            (self.order_combo, 165),
        )
        for combo, base_max in caps:
            combo.setMaximumWidth(scale_px(base_max))
            combo.updateGeometry()
        if not getattr(self, "_user_resized_cols", False):
            self._apply_scaled_table_columns()
        self._schedule_row_height_adjustment()

    def _on_header_section_clicked(self, logical_index):
        if not hasattr(self, 'order_combo'):
            return
        self.order_combo.blockSignals(True)
        if logical_index == 10:  # Agenda Item
            idx = self.order_combo.findData("AGENDA")
            if idx >= 0:
                self.order_combo.setCurrentIndex(idx)
            self.proxy.setSortMode("AGENDA")
        elif logical_index == 1:  # TDoc
            idx = self.order_combo.findData("TDOC")
            if idx >= 0:
                self.order_combo.setCurrentIndex(idx)
            self.proxy.setSortMode("TDOC")
        elif logical_index == 3:  # Source / Company
            idx = self.order_combo.findData("COMPANY")
            if idx >= 0:
                self.order_combo.setCurrentIndex(idx)
            self.proxy.setSortMode("COMPANY")
        elif logical_index == 13:  # Emails
            idx = self.order_combo.findData("EMAILS")
            if idx >= 0:
                self.order_combo.setCurrentIndex(idx)
            self.proxy.setSortMode("EMAILS")
        else:
            self.proxy.setSortMode("OTHER")
        self.order_combo.blockSignals(False)

    def _copy_table_selection(self):
        indexes = sorted(self.table.selectionModel().selectedIndexes(), key=lambda x: (x.row(), x.column()))
        if not indexes:
            return
        lines, current_line, current_row = [], [], indexes[0].row()
        for idx in indexes:
            if idx.row() != current_row:
                lines.append("\t".join(current_line))
                current_line = []
                current_row = idx.row()
            cell_text = str(idx.data(Qt.UserRole) or "").strip()
            if not cell_text:
                cell_text = str(idx.data(Qt.DisplayRole) or "").strip()
            current_line.append(cell_text)
        lines.append("\t".join(current_line))
        QApplication.clipboard().setText("\n".join(lines))
        QToolTip.showText(QCursor.pos(), "📋 Copied to clipboard!", self.table)
