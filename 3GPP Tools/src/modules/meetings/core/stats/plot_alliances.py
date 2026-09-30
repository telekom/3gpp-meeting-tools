# --- File: src/modules/meetings/core/stats/plot_alliances.py ---
import pandas as pd
import plotly.graph_objects as go
import networkx as nx
import textwrap
from .stats_html import dataframe_to_html_table


def _get_cluster_letter(index: int) -> str:
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if index < 26: return alphabet[index]
    return f"{alphabet[index // 26 - 1]}{alphabet[index % 26]}"


def compute_global_communities(df, resolution):
    G = nx.Graph()
    for companies in df['Clean_Companies']:
        if len(companies) > 1:
            for i in range(len(companies)):
                for j in range(i + 1, len(companies)):
                    c1, c2 = companies[i], companies[j]
                    if G.has_edge(c1, c2):
                        G[c1][c2]['weight'] += 1
                    else:
                        G.add_edge(c1, c2, weight=1)

    if len(G.nodes) == 0:
        return {}, {}, {}, G

    communities = list(nx.community.louvain_communities(G, seed=42, resolution=resolution))
    communities.sort(key=len, reverse=True)

    community_map = {}
    cluster_names = {}
    faction_members_dict = {}

    for i, comm in enumerate(communities):
        cluster_name = f"Cluster {_get_cluster_letter(i)}"
        cluster_names[i] = cluster_name
        faction_members_dict[cluster_name] = sorted(list(comm))
        for node in comm:
            community_map[node] = i

    return community_map, cluster_names, faction_members_dict, G


def generate_alliance_plots(df, export_dir, threshold, cluster_palette, global_factions, prefix_id="Global",
                            save_html=False):
    community_map, cluster_names, faction_members_dict, master_G = global_factions

    G = nx.Graph()
    for companies in df['Clean_Companies']:
        if len(companies) > 1:
            for i in range(len(companies)):
                for j in range(i + 1, len(companies)):
                    c1, c2 = companies[i], companies[j]
                    if G.has_edge(c1, c2):
                        G[c1][c2]['weight'] += 1
                    else:
                        G.add_edge(c1, c2, weight=1)

    edges_to_remove = [(u, v) for u, v, data in G.edges(data=True) if data['weight'] < threshold]
    G.remove_edges_from(edges_to_remove)
    G.remove_nodes_from(list(nx.isolates(G)))

    html_network = "<p style='padding:20px; color:#666;'>Not enough co-signed documents to generate network graph for this view.</p>"
    html_cluster_contribs = ""
    html_cohesion_plot = ""
    html_faction_list = ""
    tables = {"network": "", "cluster": "", "cohesion": ""}

    cluster_color_map = {name: cluster_palette[i % len(cluster_palette)] for i, name in cluster_names.items()}

    if len(G.nodes) > 0 and community_map:
        html_faction_list = "<h3 style='margin-bottom: 10px; color: #333;'>Faction Membership Roster</h3><div class='factions-container'>"
        for c_idx, c_name in cluster_names.items():
            members = faction_members_dict.get(c_name, [])
            local_members = [m for m in members if m in G.nodes]
            if not local_members: continue
            member_str = ", ".join(local_members)
            box_color = cluster_color_map[c_name]
            html_faction_list += f"<div class='faction-box' style='border-left: 4px solid {box_color};'><h4>{c_name} ({len(local_members)} Active)</h4><p>{member_str}</p></div>"
        html_faction_list += "</div>"

        # Keep a stable global layout so switching AI scopes does not move companies around.
        pos = nx.spring_layout(master_G, k=0.5, seed=42, weight='weight')
        for node in G.nodes:
            if node not in pos:
                pos[node] = [0, 0]

        # Contribution volume is a more intuitive node-size metric than raw graph degree.
        node_tdoc_counts = {node: 0 for node in G.nodes}
        node_ai_counts = {node: {} for node in G.nodes}
        for _, row in df.iterrows():
            companies = row.get('Clean_Companies', [])
            if not isinstance(companies, list) or len(companies) <= 1:
                continue
            ai_label = str(row.get('AI_Display', row.get('Agenda Item', '')) or '').strip()
            for company in set(companies):
                if company not in node_tdoc_counts:
                    continue
                node_tdoc_counts[company] += 1
                if ai_label:
                    node_ai_counts[company][ai_label] = node_ai_counts[company].get(ai_label, 0) + 1

        traces = []
        max_weight = max((data['weight'] for _, _, data in G.edges(data=True)), default=1)
        edge_weights = sorted({data['weight'] for _, _, data in G.edges(data=True)})
        edge_trace_indices = []
        edge_trace_weights = []

        # One trace per edge weight lets Plotly's in-chart threshold control hide weak ties instantly.
        for weight in edge_weights:
            edge_x, edge_y = [], []
            for u, v, data in G.edges(data=True):
                if data['weight'] == weight:
                    edge_x.extend([pos[u][0], pos[v][0], None])
                    edge_y.extend([pos[u][1], pos[v][1], None])
            strength = weight / max_weight if max_weight else 0
            calc_width = 0.35 + strength * 3.25
            opacity = 0.10 + strength * 0.55
            edge_trace_indices.append(len(traces))
            edge_trace_weights.append(weight)
            traces.append(go.Scatter(
                x=edge_x, y=edge_y, mode='lines', hoverinfo='skip',
                line=dict(width=calc_width, color=f'rgba(100,116,139,{opacity:.3f})'),
                showlegend=False, name=f'{weight} shared TDocs',
                visible=(weight >= threshold),
            ))

        # Invisible midpoint markers retain precise relationship hover information.
        mid_x, mid_y, mid_text, mid_custom = [], [], [], []
        for u, v, data in G.edges(data=True):
            mid_x.append((pos[u][0] + pos[v][0]) / 2)
            mid_y.append((pos[u][1] + pos[v][1]) / 2)
            mid_text.append(f"<b>{u}</b> 🤝 <b>{v}</b><br>Shared TDocs: {data['weight']}")
            mid_custom.append(data['weight'])
        midpoint_trace_index = len(traces)
        traces.append(go.Scatter(
            x=mid_x, y=mid_y, mode='markers', hovertext=mid_text, customdata=mid_custom,
            hovertemplate='%{hovertext}<extra></extra>',
            marker=dict(size=14, color='rgba(255,255,255,0.01)', line=dict(width=0)),
            showlegend=False, name='Connections'
        ))

        # Label only the most active companies; all companies remain discoverable on hover.
        label_limit = 18
        label_nodes = set(sorted(node_tdoc_counts, key=node_tdoc_counts.get, reverse=True)[:label_limit])
        max_node_volume = max(node_tdoc_counts.values(), default=1)

        # Separate node traces by faction to provide a compact clickable Plotly legend.
        for c_idx, c_name in cluster_names.items():
            faction_nodes = [n for n in G.nodes if community_map.get(n) == c_idx]
            if not faction_nodes:
                continue
            node_x, node_y, node_text, node_size, labels = [], [], [], [], []
            for node in faction_nodes:
                node_x.append(pos[node][0])
                node_y.append(pos[node][1])
                volume = node_tdoc_counts.get(node, 0)
                node_size.append(10 + 34 * ((volume / max_node_volume) ** 0.55 if max_node_volume else 0))
                labels.append(node if node in label_nodes else '')

                neighbors = list(G.neighbors(node))
                neighbor_weights = sorted(
                    [(n, G[node][n]['weight']) for n in neighbors], key=lambda x: x[1], reverse=True
                )
                hover_info = (
                    f"<b>{node}</b><br>Faction: {c_name}<br>Co-signed TDocs: {volume}"
                    f"<br>Partners: {len(neighbors)}<br><br><b>Top Partners:</b><br>"
                )
                for neighbor, weight in neighbor_weights[:8]:
                    hover_info += f"• {neighbor} ({weight} shared)<br>"
                top_ais = sorted(node_ai_counts.get(node, {}).items(), key=lambda x: x[1], reverse=True)[:5]
                if top_ais:
                    hover_info += '<br><b>Top Agenda Items:</b><br>'
                    for ai_label, count in top_ais:
                        hover_info += f"• {ai_label} ({count})<br>"
                node_text.append(hover_info)

            traces.append(go.Scatter(
                x=node_x, y=node_y, mode='markers+text', text=labels,
                textposition='top center', hovertext=node_text,
                hovertemplate='%{hovertext}<extra></extra>', name=c_name,
                marker=dict(
                    showscale=False, size=node_size,
                    color=cluster_color_map.get(c_name, '#CCCCCC'),
                    line=dict(width=1, color='#FFFFFF')
                )
            ))

        edge_rows = []
        for u, v, data in sorted(G.edges(data=True), key=lambda e: e[2]['weight'], reverse=True):
            edge_rows.append({
                'Company A': u, 'Company B': v, 'Shared TDocs': data['weight'],
                'Faction A': cluster_names.get(community_map.get(u, -1), 'Unknown'),
                'Faction B': cluster_names.get(community_map.get(v, -1), 'Unknown'),
            })
        tables['network'] = dataframe_to_html_table(pd.DataFrame(edge_rows), 'No alliance edge data available.')

        # Useful threshold presets based on the actual edge distribution.
        candidate_thresholds = sorted({v for v in {threshold, 2, 3, 5, 10} if threshold <= v <= max_weight})
        if threshold not in candidate_thresholds:
            candidate_thresholds.append(threshold)
            candidate_thresholds.sort()
        threshold_buttons = []
        for min_weight in candidate_thresholds:
            visibility = [True] * len(traces)
            for trace_idx, trace_weight in zip(edge_trace_indices, edge_trace_weights):
                visibility[trace_idx] = trace_weight >= min_weight
            # Keep midpoint hover trace available; its markers are invisible and don't add visual clutter.
            visibility[midpoint_trace_index] = True
            threshold_buttons.append(dict(
                label=str(min_weight), method='update',
                args=[{'visible': visibility}, {'title': f'Strategic Co-Signing Alliances (Shared TDocs ≥ {min_weight})'}]
            ))

        fig_net = go.Figure(data=traces, layout=go.Layout(
            title=f'Strategic Co-Signing Alliances (Shared TDocs ≥ {threshold})',
            showlegend=True, hovermode='closest', margin=dict(b=20, l=5, r=5, t=75),
            legend=dict(title='Faction', orientation='h', yanchor='bottom', y=1.01, xanchor='left', x=0),
            updatemenus=[dict(
                type='dropdown', direction='down', x=1.0, xanchor='right', y=1.12, yanchor='top',
                buttons=threshold_buttons, showactive=True,
                bgcolor='#F8FAFC', bordercolor='#CBD5E1', font=dict(size=11)
            )],
            annotations=[dict(
                text='Minimum shared TDocs:', x=0.80, y=1.105, xref='paper', yref='paper',
                showarrow=False, font=dict(size=11, color='#475569')
            )],
            xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
            yaxis=dict(showgrid=False, zeroline=False, showticklabels=False)
        ))

        fig_net.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        if save_html:
            fig_net.write_html(str(export_dir / f'{prefix_id}_Network_Alliances.html'))

        svg_config_net = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_Faction_Network'}}
        html_network = fig_net.to_html(full_html=False, include_plotlyjs=False,
                                       default_height='100%', default_width='100%', config=svg_config_net)

        cluster_tdoc_counts = {name: 0 for name in cluster_names.values()}
        for companies in df['Clean_Companies']:
            tdoc_clusters = set()
            for comp in companies:
                if comp in community_map:
                    tdoc_clusters.add(cluster_names[community_map[comp]])
            for c_name in tdoc_clusters:
                cluster_tdoc_counts[c_name] += 1

        plot_data = []
        for c_idx, c_name in cluster_names.items():
            members_list = faction_members_dict.get(c_name, [])
            local_members = [m for m in members_list if m in G.nodes]
            if not local_members: continue

            subgraph = G.subgraph(local_members)
            internal_weight = sum([data['weight'] for u, v, data in subgraph.edges(data=True)])
            possible_edges = (len(local_members) * (len(local_members) - 1)) / 2
            cohesion_score = internal_weight / possible_edges if possible_edges > 0 else 0
            members_str = "<br>".join(textwrap.wrap(", ".join(local_members), width=60))

            plot_data.append({
                'Faction': c_name, 'Contributions': cluster_tdoc_counts.get(c_name, 0),
                'Members': members_str, 'Member Count': len(local_members),
                'Cohesion Score': round(cohesion_score, 2)
            })

        contribs_df = pd.DataFrame(plot_data)

        if not contribs_df.empty:
            contribs_df = contribs_df.sort_values('Contributions', ascending=True)
            tables["cluster"] = dataframe_to_html_table(
                contribs_df.sort_values('Contributions', ascending=False).copy()
            )
            bar_colors = [cluster_color_map[f] for f in contribs_df['Faction']]

            fig_contribs = go.Figure(go.Bar(
                x=contribs_df['Contributions'].tolist(), y=contribs_df['Faction'].tolist(), orientation='h',
                marker=dict(color=bar_colors), hovertext=contribs_df['Members'].tolist(),
                hovertemplate="<b>%{y}</b><br>Contributions: %{x}<br><br><b>Members:</b><br>%{hovertext}<extra></extra>"
            ))
            fig_contribs.update_layout(title="Total TDoc Contributions per Faction", showlegend=False)

            fig_contribs.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
            if save_html:
                fig_contribs.write_html(str(export_dir / f"{prefix_id}_Faction_Contributions.html"))

            html_cluster_contribs = fig_contribs.to_html(full_html=False, include_plotlyjs=False, default_height="100%",
                                                         default_width="100%")

            bubble_df = contribs_df[contribs_df['Contributions'] > 0].copy()
            if not bubble_df.empty:
                bubble_colors = [cluster_color_map[f] for f in bubble_df['Faction']]
                max_contrib = float(bubble_df['Contributions'].max())

                tables["cohesion"] = dataframe_to_html_table(
                    bubble_df[['Faction', 'Member Count', 'Cohesion Score', 'Contributions', 'Members']].copy()
                )
                fig_cohesion = go.Figure(go.Scatter(
                    x=bubble_df['Member Count'].tolist(), y=bubble_df['Cohesion Score'].tolist(), mode='markers',
                    text=bubble_df['Faction'].tolist(), hovertext=bubble_df['Members'].tolist(),
                    marker=dict(
                        size=bubble_df['Contributions'].tolist(), sizemode='area',
                        sizeref=2.0 * max_contrib / (50.0 ** 2) if max_contrib > 0 else 1.0, sizemin=8,
                        color=bubble_colors, line=dict(width=1, color='#fff')
                    ),
                    hovertemplate="<b>%{text}</b><br>Faction Size: %{x} Companies<br>Internal Density: %{y}<br><br><b>Members:</b><br>%{hovertext}<extra></extra>"
                ))
                fig_cohesion.update_layout(title="Faction Cohesion vs. Size", showlegend=False,
                                           xaxis_title="Active Faction Size", yaxis_title="Cohesion Density")

                if save_html:
                    fig_cohesion.write_html(str(export_dir / f"{prefix_id}_Faction_Cohesion.html"))

                # ---> SVG CONFIG FOR COHESION PLOT
                svg_config_coh = {
                    'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_Faction_Cohesion'}}
                html_cohesion_plot = fig_cohesion.to_html(full_html=False, include_plotlyjs=False,
                                                          default_height="100%", default_width="100%",
                                                          config=svg_config_coh)

    return html_network, html_cluster_contribs, html_cohesion_plot, html_faction_list, tables