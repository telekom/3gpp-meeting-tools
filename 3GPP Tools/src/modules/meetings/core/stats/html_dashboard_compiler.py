# --- File: src/modules/meetings/core/stats/html_dashboard_compiler.py ---
def compile_view_block(
    scope_id,
    kpis,
    ai_volume_html=None,
    status_html="",
    comp_html="",
    net_html="",
    cluster_html="",
    cohesion_html="",
    list_html="",
    ai_status_html=None,
    heatmap_html=None,
    doc_type_html=None,
    ai_volume_table="",
    status_table="",
    comp_table="",
    ai_status_table="",
    heatmap_table="",
    doc_type_table="",
    network_table="",
    cluster_table="",
    cohesion_table="",
    revision_html="",
    revision_table="",
    specialization_html="",
    specialization_table="",
    is_visible=False,
    total_tdocs=None,
    total_companies=None,
):
    """Compiles an individual scope panel (Global, per-WI, or per-AI) with executive KPIs and chart cards."""
    display_style = "block" if is_visible else "none"

    # Backward compatibility guard if kpis was passed as total_tdocs integer
    if not isinstance(kpis, dict):
        total_tdocs_val = kpis
        total_companies_val = ai_volume_html
        ai_volume_html = status_html
        status_html = comp_html
        comp_html = net_html
        net_html = cluster_html
        cluster_html = cohesion_html
        cohesion_html = list_html
        list_html = ai_status_html
        ai_status_html = heatmap_html
        heatmap_html = doc_type_html
        doc_type_html = None
        kpis = {
            'total_tdocs': total_tdocs_val,
            'orig_count': total_tdocs_val,
            'rev_count': 0,
            'cr_total': 0,
            'cr_agreed_count': 0,
            'cr_agree_pct': None,
            'handling_pct': 100.0,
            'handled_count': total_tdocs_val,
            'untreated_count': 0,
            'disc_total': 0,
            'ls_total': 0,
            'total_companies': total_companies_val,
            'joint_pct': 0.0,
            'scope_dim_val': '-',
            'scope_dim_title': 'Active Items',
            'scope_dim_subtext': '',
        }

    # Extract structured KPIs
    t_tdocs = kpis.get('total_tdocs', 0)
    orig_cnt = kpis.get('orig_count', 0)
    rev_cnt = kpis.get('rev_count', 0)

    cr_tot = kpis.get('cr_total', 0)
    cr_agr_cnt = kpis.get('cr_agreed_count', 0)
    cr_agr_pct = kpis.get('cr_agree_pct')
    cr_agr_str = f"{cr_agr_pct}%" if cr_agr_pct is not None else "N/A"
    cr_sub = f"{cr_agr_cnt} of {cr_tot} CRs Agreed" if cr_tot > 0 else "No CRs in this scope"

    h_pct = kpis.get('handling_pct', 0.0)
    h_cnt = kpis.get('handled_count', 0)
    u_cnt = kpis.get('untreated_count', 0)

    disc_tot = kpis.get('disc_total', 0)
    ls_tot = kpis.get('ls_total', 0)
    cr_pct = round((cr_tot / t_tdocs) * 100) if t_tdocs else 0
    disc_pct = round((disc_tot / t_tdocs) * 100) if t_tdocs else 0
    doc_ratio_val = f"{cr_tot} / {disc_tot} / {ls_tot}"
    doc_ratio_sub = f"{cr_pct}% CR · {disc_pct}% Disc"

    t_comps = kpis.get('total_companies', 0)
    j_pct = kpis.get('joint_pct', 0.0)

    s_val = kpis.get('scope_dim_val', '-')
    s_title = kpis.get('scope_dim_title', 'Scope Items')
    s_sub = kpis.get('scope_dim_subtext', '')

    kpi_strip = f"""
    <div class="kpi-container">
        <div class="kpi-card">
            <h3>{t_tdocs}</h3>
            <p>Total TDocs</p>
            <div class="kpi-subtext">{orig_cnt} Originals · {rev_cnt} Revisions</div>
        </div>
        <div class="kpi-card">
            <h3>{cr_agr_str}</h3>
            <p>CR Agreement Rate</p>
            <div class="kpi-subtext">{cr_sub}</div>
        </div>
        <div class="kpi-card">
            <h3>{h_pct}%</h3>
            <p>Treatment Rate</p>
            <div class="kpi-subtext">{h_cnt} Handled · {u_cnt} Untreated</div>
        </div>
        <div class="kpi-card">
            <h3>{doc_ratio_val}</h3>
            <p>CR / Disc / LS</p>
            <div class="kpi-subtext">{doc_ratio_sub}</div>
        </div>
        <div class="kpi-card">
            <h3>{t_comps}</h3>
            <p>Active Entities</p>
            <div class="kpi-subtext">{j_pct}% Joint Contributions</div>
        </div>
        <div class="kpi-card">
            <h3>{s_val}</h3>
            <p>{s_title}</p>
            <div class="kpi-subtext">{s_sub}</div>
        </div>
    </div>
    """

    volume_card = ""
    if ai_volume_html:
        volume_card = """
        <div class="chart-card">
            <div class="info-title-container">
                <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Agenda Items by Volume:</b> Displays top topics based on submitted document frequency.</span></span>
            </div>
            <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
            <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
            <div class="chart-pane">__AI_VOLUME_HTML__</div><div class="data-pane">__AI_VOLUME_TABLE__</div>
        </div>
        """.replace("__AI_VOLUME_HTML__", str(ai_volume_html)).replace("__AI_VOLUME_TABLE__", str(ai_volume_table))

    doc_type_card = ""
    if doc_type_html:
        doc_card_span = 'style="grid-column: 1 / -1; height: 480px;"' if ai_volume_html else 'style="grid-column: span 1;"'
        doc_type_card = f"""
        <div class="chart-card" {doc_card_span}>
            <div class="info-title-container">
                <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Document Types by Outcome:</b> Breakdown of Change Requests, Discussion Papers, and Liaison Statements segmented by meeting status.</span></span>
            </div>
            <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
            <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
            <div class="chart-pane">__DOC_TYPE_HTML__</div><div class="data-pane">__DOC_TYPE_TABLE__</div>
        </div>
        """.replace("__DOC_TYPE_HTML__", str(doc_type_html)).replace("__DOC_TYPE_TABLE__", str(doc_type_table))

    ai_status_card = ""
    if ai_status_html:
        ai_status_card = """
        <div class="chart-card" style="grid-column: 1 / -1; height: 500px;">
            <div class="info-title-container">
                <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Agenda Items by Status:</b> Breakdown of outcomes across the top topics.</span></span>
            </div>
            <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
            <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
            <div class="chart-pane">__AI_STATUS_HTML__</div><div class="data-pane">__AI_STATUS_TABLE__</div>
        </div>
        """.replace("__AI_STATUS_HTML__", str(ai_status_html)).replace("__AI_STATUS_TABLE__", str(ai_status_table))

    heatmap_card = ""
    if heatmap_html:
        heatmap_card = """
        <div class="chart-card" style="grid-column: 1 / -1; height: 650px;">
            <div class="info-title-container">
                <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Company Focus Matrix:</b> Heatmap of TDoc submissions by the top companies across top topics.</span></span>
            </div>
            <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
            <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
            <div class="chart-pane">__HEATMAP_HTML__</div><div class="data-pane">__HEATMAP_TABLE__</div>
        </div>
        """.replace("__HEATMAP_HTML__", str(heatmap_html)).replace("__HEATMAP_TABLE__", str(heatmap_table))

    revision_card = ""
    if revision_html:
        revision_card = """
        <div class="chart-card" style="grid-column: 1 / -1; height: 520px;">
            <div class="info-title-container"><span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Revision Activity:</b> Shows original TDocs and subsequent revisions per Agenda Item. Revision intensity is descriptive and does not imply quality.</span></span></div>
            <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
            <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
            <div class="chart-pane">__REVISION_HTML__</div><div class="data-pane">__REVISION_TABLE__</div>
        </div>
        """.replace("__REVISION_HTML__", str(revision_html)).replace("__REVISION_TABLE__", str(revision_table))

    specialization_card = ""
    if specialization_html:
        specialization_card = """
        <div class="chart-card" style="grid-column: 1 / -1; height: 560px;">
            <div class="info-title-container"><span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Company Topic Specialization:</b> Bubble size is contribution volume; horizontal position is the share of each company's TDocs concentrated in its most frequent topic.</span></span></div>
            <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
            <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
            <div class="chart-pane">__SPECIALIZATION_HTML__</div><div class="data-pane">__SPECIALIZATION_TABLE__</div>
        </div>
        """.replace("__SPECIALIZATION_HTML__", str(specialization_html)).replace("__SPECIALIZATION_TABLE__", str(specialization_table))

    grid_col_span = "grid-column: span 1;"

    html_template = f"""
    <div id="__SCOPE_ID__" class="dashboard-view-panel" style="display: __DISPLAY_STYLE__;">
        {kpi_strip}

        <div class="grid-container">
            __VOLUME_CARD__
            <div class="chart-card" style="__GRID_COL_SPAN__">
                <div class="info-title-container">
                    <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>TDoc Outcomes:</b> Breakdown of actions applied across this data selection.</span></span>
                </div>
                <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                <div class="chart-pane">__STATUS_HTML__</div><div class="data-pane">__STATUS_TABLE__</div>
            </div>

            __DOC_TYPE_CARD__

            __AI_STATUS_CARD__
            __HEATMAP_CARD__

            <div class="chart-card" style="grid-column: 1 / -1; height: 550px;">
                <div class="info-title-container">
                    <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Top Contributors:</b> Active entities in the scope subset.</span></span>
                </div>
                <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                <div class="chart-pane">__COMP_HTML__</div><div class="data-pane">__COMP_TABLE__</div>
            </div>

            <div class="chart-card" style="grid-column: 1 / -1; height: 750px;">
                <div class="info-title-container">
                    <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Strategic Alliances:</b> Subset collaboration mappings under fixed global configurations.</span></span>
                </div>
                <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                <div class="chart-pane">__NET_HTML__</div><div class="data-pane">__NETWORK_TABLE__</div>
            </div>

            <div class="chart-card" style="height: 450px;">
                <div class="info-title-container">
                    <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Faction Output Volume:</b> Quantifies documents matching this view scope authored by the tracked alliance clusters.</span></span>
                </div>
                <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                <div class="chart-pane">__CLUSTER_HTML__</div><div class="data-pane">__CLUSTER_TABLE__</div>
            </div>

            <div class="chart-card" style="height: 450px;">
                <div class="info-title-container">
                    <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Cohesion Tracker:</b> Displays the interaction density of factions under this specific subset.</span></span>
                </div>
                <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                <div class="chart-pane">__COHESION_HTML__</div><div class="data-pane">__COHESION_TABLE__</div>
            </div>

            __REVISION_CARD__
            __SPECIALIZATION_CARD__

            <div class="chart-card" style="grid-column: 1 / -1; height: auto; padding: 20px;">
                __LIST_HTML__
            </div>
        </div>
    </div>
    """

    html_template = html_template.replace("__SCOPE_ID__", str(scope_id))
    html_template = html_template.replace("__DISPLAY_STYLE__", str(display_style))
    html_template = html_template.replace("__VOLUME_CARD__", str(volume_card))
    html_template = html_template.replace("__GRID_COL_SPAN__", str(grid_col_span))
    html_template = html_template.replace("__STATUS_HTML__", str(status_html))
    html_template = html_template.replace("__STATUS_TABLE__", str(status_table))
    html_template = html_template.replace("__DOC_TYPE_CARD__", str(doc_type_card))

    html_template = html_template.replace("__AI_STATUS_CARD__", str(ai_status_card))
    html_template = html_template.replace("__HEATMAP_CARD__", str(heatmap_card))

    html_template = html_template.replace("__COMP_HTML__", str(comp_html))
    html_template = html_template.replace("__COMP_TABLE__", str(comp_table))
    html_template = html_template.replace("__NET_HTML__", str(net_html))
    html_template = html_template.replace("__NETWORK_TABLE__", str(network_table))
    html_template = html_template.replace("__CLUSTER_HTML__", str(cluster_html))
    html_template = html_template.replace("__CLUSTER_TABLE__", str(cluster_table))
    html_template = html_template.replace("__COHESION_HTML__", str(cohesion_html))
    html_template = html_template.replace("__COHESION_TABLE__", str(cohesion_table))
    html_template = html_template.replace("__REVISION_CARD__", str(revision_card))
    html_template = html_template.replace("__SPECIALIZATION_CARD__", str(specialization_card))
    html_template = html_template.replace("__LIST_HTML__", str(list_html))

    return html_template
