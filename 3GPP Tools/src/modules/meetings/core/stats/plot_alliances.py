# --- File: src/modules/meetings/core/stats/plot_alliances.py ---
import pandas as pd
import plotly.graph_objects as go
import networkx as nx
import textwrap


def _get_cluster_letter(index: int) -> str:
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if index < 26: return alphabet[index]
    return f"{alphabet[index // 26 - 1]}{alphabet[index % 26]}"


def generate_alliance_plots(df, export_dir, threshold, resolution, cluster_palette):
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

    html_net = "<p style='padding:20px; color:#666;'>Not enough co-signed documents to generate network graph.</p>"
    html_cluster_contribs = ""
    html_cohesion_plot = ""
    html_faction_list = ""

    if len(G.nodes) > 0:
        communities = list(nx.community.louvain_communities(G, seed=42, resolution=resolution))
        communities.sort(key=len, reverse=True)

        community_map = {}
        cluster_names = {}
        faction_members_dict = {}
        cluster_color_map = {}

        for i, comm in enumerate(communities):
            cluster_name = f"Cluster {_get_cluster_letter(i)}"
            cluster_names[i] = cluster_name
            cluster_color_map[cluster_name] = cluster_palette[i % len(cluster_palette)]
            faction_members_dict[cluster_name] = sorted(list(comm))
            for node in comm:
                community_map[node] = i

        html_faction_list = "<h3 style='margin-bottom: 10px; color: #333;'>Faction Membership Roster</h3><div class='factions-container'>"
        for c_name, members in faction_members_dict.items():
            member_str = ", ".join(members)
            box_color = cluster_color_map[c_name]
            html_faction_list += f"<div class='faction-box' style='border-left: 4px solid {box_color};'><h4>{c_name} ({len(members)} Members)</h4><p>{member_str}</p></div>"
        html_faction_list += "</div>"

        # --- 1. Network Graph ---
        pos = nx.spring_layout(G, k=0.5, seed=42)
        traces = []

        max_weight = max([data['weight'] for u, v, data in G.edges(data=True)]) if G.edges else 1
        edge_weights = set([data['weight'] for u, v, data in G.edges(data=True)])

        # Visual Lines
        for weight in edge_weights:
            edge_x, edge_y = [], []
            for u, v, data in G.edges(data=True):
                if data['weight'] == weight:
                    edge_x.extend([pos[u][0], pos[v][0], None])
                    edge_y.extend([pos[u][1], pos[v][1], None])

            calc_width = 0.5 + (weight / max_weight) * 4.0
            traces.append(go.Scatter(x=edge_x, y=edge_y, line=dict(width=calc_width, color='#999'),
                                     hoverinfo='none', mode='lines', opacity=0.6))

        # Invisible Midpoints (Hitboxes for Edge Hovertext)
        mid_x, mid_y, mid_text = [], [], []
        for u, v, data in G.edges(data=True):
            x0, y0 = pos[u]
            x1, y1 = pos[v]
            mid_x.append((x0 + x1) / 2)
            mid_y.append((y0 + y1) / 2)
            mid_text.append(f"<b>{u}</b> 🤝 <b>{v}</b><br>Shared TDocs: {data['weight']}")

        traces.append(go.Scatter(
            x=mid_x, y=mid_y, mode='markers',
            hovertext=mid_text,
            hovertemplate="%{hovertext}<extra></extra>",
            marker=dict(size=14, color='rgba(255,255,255,0.01)', line=dict(width=0)),
            showlegend=False, name="Connections"
        ))

        # Visual Nodes
        node_x, node_y, node_text, node_size, node_color = [], [], [], [], []
        for node in G.nodes():
            x, y = pos[node]
            node_x.append(x)
            node_y.append(y)

            adj_size = len(list(G.neighbors(node))) * 3
            node_size.append(max(10, min(adj_size, 50)))

            c_name = cluster_names[community_map[node]]
            node_color.append(cluster_color_map[c_name])

            neighbors = list(G.neighbors(node))

            hover_info = f"<b>{node}</b><br>Faction: {c_name}<br>Partners: {len(neighbors)}<br><br><b>Top Partners:</b><br>"
            neighbor_weights = [(n, G[node][n]['weight']) for n in neighbors]
            neighbor_weights.sort(key=lambda item: item[1], reverse=True)

            for neighbor, weight in neighbor_weights[:15]:
                hover_info += f"• {neighbor} ({weight} shared)<br>"
            if len(neighbors) > 15: hover_info += f"<i>...and {len(neighbors) - 15} more</i>"

            node_text.append(hover_info)

        traces.append(go.Scatter(
            x=node_x, y=node_y, mode='markers+text', text=list(G.nodes()),
            textposition="top center",
            hovertext=node_text,
            hovertemplate="%{hovertext}<extra></extra>",
            name="Companies",
            marker=dict(showscale=False, size=node_size, color=node_color,
                        line_width=1, line_color='#fff')
        ))

        fig_net = go.Figure(data=traces,
                            layout=go.Layout(title=f'Strategic Co-Signing Alliances (Threshold >= {threshold})',
                                             showlegend=False, hovermode='closest',
                                             margin=dict(b=20, l=5, r=5, t=40),
                                             xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
                                             yaxis=dict(showgrid=False, zeroline=False, showticklabels=False)))

        fig_net.write_html(str(export_dir / "Stat_Network_Alliances.html"))
        html_net = fig_net.to_html(full_html=False, include_plotlyjs=False, div_id="network_graph",
                                   default_height="100%", default_width="100%")

        # --- 2. Calculate Cohesion & Contributions ---
        cluster_tdoc_counts = {c_name: 0 for c_name in cluster_names.values()}
        for companies in df['Clean_Companies']:
            tdoc_clusters = set()
            for comp in companies:
                if comp in community_map:
                    tdoc_clusters.add(cluster_names[community_map[comp]])
            for c_name in tdoc_clusters:
                cluster_tdoc_counts[c_name] += 1

        plot_data = []
        for c_name, count in cluster_tdoc_counts.items():
            members_list = faction_members_dict.get(c_name, [])

            subgraph = G.subgraph(members_list)
            internal_weight = sum([data['weight'] for u, v, data in subgraph.edges(data=True)])
            possible_edges = (len(members_list) * (len(members_list) - 1)) / 2

            cohesion_score = internal_weight / possible_edges if possible_edges > 0 else 0

            members_str = "<br>".join(textwrap.wrap(", ".join(members_list), width=60))
            plot_data.append({
                'Faction': c_name,
                'Contributions': count,
                'Members': members_str,
                'Member Count': len(members_list),
                'Cohesion Score': round(cohesion_score, 2)
            })

        # --- 3. Contributions Bar Chart (Direct go.Bar) ---
        contribs_df = pd.DataFrame(plot_data).sort_values('Contributions', ascending=True)
        bar_colors = [cluster_color_map[f] for f in contribs_df['Faction']]

        fig_contribs = go.Figure(go.Bar(
            x=contribs_df['Contributions'],
            y=contribs_df['Faction'],
            orientation='h',
            marker=dict(color=bar_colors),
            hovertext=contribs_df['Members'],
            hovertemplate="<b>%{y}</b><br>Contributions: %{x}<br><br><b>Members:</b><br>%{hovertext}<extra></extra>"
        ))

        fig_contribs.update_layout(title="Total TDoc Contributions per Faction (Louvain method)", showlegend=False)
        fig_contribs.write_html(str(export_dir / "Stat_Faction_Contributions.html"))
        html_cluster_contribs = fig_contribs.to_html(full_html=False, include_plotlyjs=False, default_height="100%",
                                                     default_width="100%")

        # --- 4. Cohesion Bubble Chart (Direct go.Scatter) ---
        bubble_df = contribs_df[contribs_df['Contributions'] > 0].copy()
        bubble_colors = [cluster_color_map[f] for f in bubble_df['Faction']]
        max_contrib = bubble_df['Contributions'].max() if not bubble_df.empty else 1

        fig_cohesion = go.Figure(go.Scatter(
            x=bubble_df['Member Count'],
            y=bubble_df['Cohesion Score'],
            mode='markers',
            text=bubble_df['Faction'],
            hovertext=bubble_df['Members'],
            marker=dict(
                size=bubble_df['Contributions'],
                sizemode='area',
                sizeref=2.0 * max_contrib / (50.0 ** 2) if max_contrib > 0 else 1,
                sizemin=8,
                color=bubble_colors,
                line=dict(width=1, color='#fff')
            ),
            hovertemplate="<b>%{text}</b><br>Faction Size: %{x} Companies<br>Internal Cohesion Density: %{y}<br><br><b>Members:</b><br>%{hovertext}<extra></extra>"
        ))

        fig_cohesion.update_layout(
            title="Faction Cohesion vs. Size (Bubble = Output Volume)",
            showlegend=False,
            xaxis_title="Faction Size (Number of Companies)",
            yaxis_title="Cohesion Score (Network Density)"
        )
        fig_cohesion.write_html(str(export_dir / "Stat_Faction_Cohesion.html"))
        html_cohesion_plot = fig_cohesion.to_html(full_html=False, include_plotlyjs=False, default_height="100%",
                                                  default_width="100%")

    return html_net, html_cluster_contribs, html_cohesion_plot, html_faction_list