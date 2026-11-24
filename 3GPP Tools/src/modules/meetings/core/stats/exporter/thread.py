"""QThread orchestrator generating multi-scope meeting statistics dashboards and workbooks."""

from __future__ import annotations

import html
import json
import re
from pathlib import Path

import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal
from plotly.offline import get_plotlyjs_version

from core.config.plot_styles import CLUSTER_PALETTE, PALETTE, THEME_COLOR
from modules.meetings.core.stats.exporter.html_shell import render_dashboard_html
from modules.meetings.core.stats.exporter.kpi_and_workbook import (
    compute_scope_kpis,
    export_statistics_workbook,
    prepare_statistics_dataframe,
)
from modules.meetings.core.stats.html_dashboard_compiler import compile_view_block
from modules.meetings.core.stats.plot_agenda import generate_ai_status_plot, generate_ai_volume_plot
from modules.meetings.core.stats.plot_alliances import compute_global_communities, generate_alliance_plots
from modules.meetings.core.stats.plot_contributors import (
    generate_company_ai_heatmap,
    generate_top_contributors_plot,
)
from modules.meetings.core.stats.plot_insights import (
    generate_company_specialization_plot,
    generate_revision_intensity_plot,
)
from modules.meetings.core.stats.plot_status import generate_doc_types_plot, generate_outcomes_plot


def _natural_sort_key(value):
    return [int(text) if text.isdigit() else text.lower() for text in re.split("([0-9]+)", str(value))]


class StatisticsExporterThread(QThread):
    """Background QThread compiling multi-scope Plotly HTML dashboards and Excel stats."""

    finished = pyqtSignal(bool, str)

    def __init__(self, meeting_dir: Path, tdocs_data: list, mtg_info: dict, config: dict):
        super().__init__()
        self.meeting_dir = meeting_dir
        self.tdocs_data = tdocs_data
        self.mtg_info = mtg_info
        self.config = config
        self.export_dir = self.meeting_dir / "Export"

        self.cfg_resolution = self.config.get("resolution", 1.5)
        self.cfg_threshold = self.config.get("threshold", 1)
        self.cfg_top_count = self.config.get("top_count", 30)
        self.cfg_export_html = self.config.get("export_html_plots", False)
        self.cfg_hm_companies = self.config.get("heatmap_top_companies", 25)
        self.cfg_hm_ais = self.config.get("heatmap_top_ais", 25)

        self.THEME_COLOR = THEME_COLOR
        self.PALETTE = PALETTE
        self.CLUSTER_PALETTE = CLUSTER_PALETTE

    def _compute_scope_kpis(self, scope_df: pd.DataFrame, scope_type: str = "global") -> dict:
        return compute_scope_kpis(scope_df, scope_type=scope_type)

    def run(self):
        try:
            self.export_dir.mkdir(parents=True, exist_ok=True)
            plots_dir = self.export_dir / "Interactive_Plots"
            if self.cfg_export_html:
                plots_dir.mkdir(parents=True, exist_ok=True)

            df, interactive_rows = prepare_statistics_dataframe(self.tdocs_data, self.meeting_dir)
            if df.empty:
                self.finished.emit(False, "No TDoc data available to generate statistics.")
                return

            interactive_json = json.dumps(interactive_rows, ensure_ascii=False).replace("</", "<\\/")
            global_factions = compute_global_communities(df, self.cfg_resolution)

            # Global plots
            g_html_ai, g_table_ai = generate_ai_volume_plot(
                df, plots_dir, self.THEME_COLOR, prefix_id="Global", save_html=self.cfg_export_html
            )
            g_html_ai_status, g_table_ai_status = generate_ai_status_plot(
                df, plots_dir, self.PALETTE, prefix_id="Global", save_html=self.cfg_export_html
            )
            g_html_status, g_table_status = generate_outcomes_plot(
                df, plots_dir, self.PALETTE, prefix_id="Global", save_html=self.cfg_export_html
            )
            g_html_comp, _total_companies, g_table_comp = generate_top_contributors_plot(
                df, plots_dir, self.THEME_COLOR, self.cfg_top_count, prefix_id="Global", save_html=self.cfg_export_html
            )
            g_html_heatmap, g_table_heatmap = generate_company_ai_heatmap(
                df,
                plots_dir,
                prefix_id="Global",
                save_html=self.cfg_export_html,
                top_comps_count=self.cfg_hm_companies,
                top_ais_count=self.cfg_hm_ais,
            )
            g_html_net, g_html_cluster, g_html_cohesion, g_html_list, g_alliance_tables = generate_alliance_plots(
                df,
                plots_dir,
                self.cfg_threshold,
                self.CLUSTER_PALETTE,
                global_factions,
                prefix_id="Global",
                save_html=self.cfg_export_html,
            )
            g_html_revision, g_table_revision = generate_revision_intensity_plot(
                df, plots_dir, self.THEME_COLOR, prefix_id="Global", save_html=self.cfg_export_html
            )
            g_html_specialization, g_table_specialization = generate_company_specialization_plot(
                df, plots_dir, self.THEME_COLOR, prefix_id="Global", save_html=self.cfg_export_html
            )

            raw_ais = df["Agenda Item"].dropna().unique()
            unique_ais = sorted([str(ai).strip() for ai in raw_ais if str(ai).strip()], key=_natural_sort_key)
            meeting_name = f"{self.mtg_info.get('wg_name', 'WG')} {self.mtg_info.get('meeting_number', '')}"

            views_html_buffer = []
            dropdown_options = ['<option value="global">🌐 Overall Meeting View</option>']

            g_kpis = self._compute_scope_kpis(df, "global")
            g_html_doc_type, g_table_doc_type = generate_doc_types_plot(
                df, plots_dir, prefix_id="Global", save_html=self.cfg_export_html
            )
            views_html_buffer.append(
                self._compile_view_block(
                    "global",
                    g_kpis,
                    g_html_ai,
                    g_html_status,
                    g_html_comp,
                    g_html_net,
                    g_html_cluster,
                    g_html_cohesion,
                    g_html_list,
                    ai_status_html=g_html_ai_status,
                    heatmap_html=g_html_heatmap,
                    doc_type_html=g_html_doc_type,
                    ai_volume_table=g_table_ai,
                    status_table=g_table_status,
                    comp_table=g_table_comp,
                    ai_status_table=g_table_ai_status,
                    heatmap_table=g_table_heatmap,
                    doc_type_table=g_table_doc_type,
                    network_table=g_alliance_tables.get("network", ""),
                    cluster_table=g_alliance_tables.get("cluster", ""),
                    cohesion_table=g_alliance_tables.get("cohesion", ""),
                    revision_html=g_html_revision,
                    revision_table=g_table_revision,
                    specialization_html=g_html_specialization,
                    specialization_table=g_table_specialization,
                    is_visible=True,
                )
            )

            self._append_wi_views(df, plots_dir, global_factions, dropdown_options, views_html_buffer)
            self._append_ai_views(df, plots_dir, global_factions, unique_ais, dropdown_options, views_html_buffer)

            dashboard_html = render_dashboard_html(
                meeting_name=meeting_name,
                dropdown_options=dropdown_options,
                interactive_json=interactive_json,
                views_html_buffer=views_html_buffer,
                plotly_js_version=get_plotlyjs_version(),
            )
            self._export_statistics_workbook(df, self.export_dir / "Statistics_Data.xlsx")
            out_file = self.export_dir / "Statistics_Report.html"
            with open(out_file, "w", encoding="utf-8") as handle:
                handle.write(dashboard_html)

            self.finished.emit(True, str(out_file))
        except Exception as exc:
            self.finished.emit(False, str(exc))

    def _append_wi_views(self, df, plots_dir, global_factions, dropdown_options, views_html_buffer):
        unique_wis = sorted({wi for wis in df["Clean_WI_List"] for wi in wis}, key=lambda x: x.lower())
        if unique_wis:
            dropdown_options.append("<option disabled>──────── Work / Study Items ────────</option>")
        for wi_idx, wi_name in enumerate(unique_wis):
            wi_df = df[df["Clean_WI_List"].apply(lambda values, w=wi_name: w in values)].copy().reset_index(drop=True)
            if wi_df.empty:
                continue
            safe_id = f"wi_{wi_idx}"
            safe_wi_prefix = "WI_" + re.sub(r'[\\/*?:"<>|]', "_", wi_name)
            dropdown_options.append(
                f'<option value="{safe_id}" data-wi="{html.escape(wi_name, quote=True)}">'
                f"🧩 {html.escape(wi_name)} ({len(wi_df)} TDocs)</option>"
            )
            wi_html_ai, wi_table_ai = generate_ai_volume_plot(wi_df, plots_dir, self.THEME_COLOR, safe_wi_prefix, self.cfg_export_html)
            wi_html_status, wi_table_status = generate_outcomes_plot(wi_df, plots_dir, self.PALETTE, safe_wi_prefix, self.cfg_export_html)
            wi_html_comp, _wi_companies, wi_table_comp = generate_top_contributors_plot(
                wi_df, plots_dir, self.THEME_COLOR, self.cfg_top_count, safe_wi_prefix, self.cfg_export_html
            )
            wi_html_net, wi_html_cluster, wi_html_cohesion, wi_html_list, wi_alliance_tables = generate_alliance_plots(
                wi_df, plots_dir, self.cfg_threshold, self.CLUSTER_PALETTE, global_factions, safe_wi_prefix, self.cfg_export_html
            )
            wi_html_revision, wi_table_revision = generate_revision_intensity_plot(
                wi_df, plots_dir, self.THEME_COLOR, safe_wi_prefix, self.cfg_export_html
            )
            wi_kpis = self._compute_scope_kpis(wi_df, "wi")
            wi_html_doc_type, wi_table_doc_type = generate_doc_types_plot(
                wi_df, plots_dir, prefix_id=safe_wi_prefix, save_html=self.cfg_export_html
            )
            views_html_buffer.append(
                self._compile_view_block(
                    safe_id,
                    wi_kpis,
                    wi_html_ai,
                    wi_html_status,
                    wi_html_comp,
                    wi_html_net,
                    wi_html_cluster,
                    wi_html_cohesion,
                    wi_html_list,
                    doc_type_html=wi_html_doc_type,
                    ai_volume_table=wi_table_ai,
                    status_table=wi_table_status,
                    comp_table=wi_table_comp,
                    doc_type_table=wi_table_doc_type,
                    network_table=wi_alliance_tables.get("network", ""),
                    cluster_table=wi_alliance_tables.get("cluster", ""),
                    cohesion_table=wi_alliance_tables.get("cohesion", ""),
                    revision_html=wi_html_revision,
                    revision_table=wi_table_revision,
                    is_visible=False,
                )
            )

    def _append_ai_views(self, df, plots_dir, global_factions, unique_ais, dropdown_options, views_html_buffer):
        if unique_ais:
            dropdown_options.append("<option disabled>──────── Agenda Items ────────</option>")
        for idx, ai_name in enumerate(unique_ais):
            ai_df = df[df["Agenda Item"].astype(str).str.strip() == ai_name].copy().reset_index(drop=True)
            if ai_df.empty:
                continue

            safe_id = f"ai_{idx}"
            clean_ai_name = re.sub(r'[\\/*?:"<>|]', "_", str(ai_name))
            safe_ai_prefix = "AI_" + clean_ai_name

            ai_display = str(ai_df["AI_Display"].iloc[0]) if "AI_Display" in ai_df.columns else ai_name
            dropdown_options.append(
                f'<option value="{safe_id}" data-ai="{html.escape(ai_name, quote=True)}">'
                f"📌 {html.escape(ai_display)} ({len(ai_df)} TDocs)</option>"
            )

            ai_html_status, ai_table_status = generate_outcomes_plot(
                ai_df, plots_dir, self.PALETTE, safe_ai_prefix, save_html=self.cfg_export_html
            )
            ai_html_comp, _ai_companies, ai_table_comp = generate_top_contributors_plot(
                ai_df, plots_dir, self.THEME_COLOR, self.cfg_top_count, safe_ai_prefix, save_html=self.cfg_export_html
            )
            ai_html_net, ai_html_cluster, ai_html_cohesion, ai_html_list, ai_alliance_tables = generate_alliance_plots(
                ai_df,
                plots_dir,
                self.cfg_threshold,
                self.CLUSTER_PALETTE,
                global_factions,
                safe_ai_prefix,
                save_html=self.cfg_export_html,
            )

            ai_kpis = self._compute_scope_kpis(ai_df, "ai")
            ai_html_doc_type, ai_table_doc_type = generate_doc_types_plot(
                ai_df, plots_dir, prefix_id=safe_ai_prefix, save_html=self.cfg_export_html
            )
            # Keep AI Status and Heatmap explicitly None for individual Agenda Items
            views_html_buffer.append(
                self._compile_view_block(
                    safe_id,
                    ai_kpis,
                    ai_volume_html=None,
                    status_html=ai_html_status,
                    comp_html=ai_html_comp,
                    net_html=ai_html_net,
                    cluster_html=ai_html_cluster,
                    cohesion_html=ai_html_cohesion,
                    list_html=ai_html_list,
                    ai_status_html=None,
                    heatmap_html=None,
                    doc_type_html=ai_html_doc_type,
                    status_table=ai_table_status,
                    comp_table=ai_table_comp,
                    doc_type_table=ai_table_doc_type,
                    network_table=ai_alliance_tables.get("network", ""),
                    cluster_table=ai_alliance_tables.get("cluster", ""),
                    cohesion_table=ai_alliance_tables.get("cohesion", ""),
                    is_visible=False,
                )
            )

    def _export_statistics_workbook(self, df: pd.DataFrame, path: Path):
        export_statistics_workbook(df, path)

    def _compile_view_block(self, *args, **kwargs) -> str:
        return compile_view_block(*args, **kwargs)
