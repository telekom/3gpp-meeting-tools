# --- File: src/modules/meetings/core/stats/plot_alliances.py ---
import pandas as pd
import plotly.graph_objects as go
import networkx as nx
import textwrap


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


def generate_alliance_plots(df, export_dir, threshold, cluster_palette, global_factions, prefix_id="Global"):
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

    html_net = "<p style='padding:20px; color:#666;'>Not enough co-signed documents to generate network graph for this view.</p>"
    html_cluster_contribs = ""
    html_cohesion_plot = ""
    html_faction_list = ""

    cluster_color_map = {name: cluster_palette[i % len(cluster_palette)] for i, name in cluster_names.items()}

    if len(G.nodes) > 0 and community_map:
        html_faction_list = "<h3 style='margin-bottom: 10px; color: #333;'>Faction Membership Roster</h3><div class='factions-container'>"
        for c_idx, c_name in cluster_names.items():
            members = faction_members_dict.get(c_name, [])
            local_members = [m for members in members if (m := members) in G.nodes]
            if not local_members: continue
            member_str = ", ".join(local_members)
            box_color = cluster_color_map[c_name]
            html_faction_list += f"<div class='faction-box' style='border-left: 4px solid {box_color};'><h4>{c_name} ({len(local_members)} Active)</h4><p>{member_str}</p></div>"
        html_faction_list += "</div>"

        pos = nx.spring_layout(master_G, k=0.5, seed=42)
        for node in G.nodes:
            if node not in pos: pos[node] = [0, 0]

        traces = []
        max_weight = max([data['weight'] for u, v, data in G.edges(data=True)]) if G.edges else 1
        edge_weights = set([data['weight'] for u, v, data in G.edges(data=True)])

        for weight in edge_weights:
            edge_x, edge_y = [], []
            for u, v, data in G.edges(data=True):
                if data['weight'] == weight:
                    edge_x.extend([pos[u][0], pos[v][0], None])
                    edge_y.extend([pos[u][1], pos[v][1], None])
            calc_width = 0.5 + (weight / max_weight) * 4.0
            traces.append(go.Scatter(x=edge_x, y=edge_y, line=dict(width=calc_width, color='#999'),
                                     hoverinfo='none', mode='lines', opacity=0.6))

        mid_x, mid_y, mid_text = [], [], []
        for u, v, data in G.edges(data=True):
            mid_x.append((pos[u][0] + pos[v][0]) / 2)
            mid_y.append((pos[u][1] + pos[v][1]) / 2)
            mid_text.append(f"<b>{u}</b> 🤝 <b>{v}</b><br>Shared TDocs: {data['weight']}")

        traces.append(go.Scatter(
            x=mid_x, y=mid_y, mode='markers', hovertext=mid_text,
            hovertemplate="%{hovertext}<extra></extra>",
            marker=dict(size=14, color='rgba(255,255,255,0.01)', line=dict(width=0)),
            showlegend=False, name="Connections"
        ))

        node_x, node_y, node_text, node_size, node_color = [], [], [], [], []
        for node in G.nodes():
            node_x.append(pos[node][0])
            node_y.append(pos[node][1])
            node_size.append(max(10, min(len(list(G.neighbors(node))) * 3, 50)))

            c_idx = community_map.get(node, 0)
            c_name = cluster_names.get(c_idx, "Unknown")
            node_color.append(cluster_color_map.get(c_name, "#CCCCCC"))

            neighbors = list(G.neighbors(node))
            hover_info = f"<b>{node}</b><br>Faction: {c_name}<br>Partners: {len(neighbors)}<br><br><b>Top Partners:</b><br>"
            neighbor_weights = sorted([(n, G[node][n]['weight']) for n in neighbors], key=lambda x: x[1], reverse=True)

            for neighbor, weight in neighbor_weights[:15]:
                hover_info += f"• {neighbor} ({weight} shared)<br>"
            node_text.append(hover_info)

        traces.append(go.Scatter(
            x=node_x, y=node_y, mode='markers+text', text=list(G.nodes()),
            textposition="top center", hovertext=node_text,
            hovertemplate="%{hovertext}<extra></extra>", name="Companies",
            marker=dict(showscale=False, size=node_size, color=node_color, line_width=1, line_color='#fff')
        ))

        fig_net = go.Figure(data=traces, layout=go.Layout(
            title=f'Strategic Co-Signing Alliances (Threshold >= {threshold})',
            showlegend=False, hovermode='closest', margin=dict(b=20, l=5, r=5, t=40),
            xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
            yaxis=dict(showgrid=False, zeroline=False, showticklabels=False)
        ))

        # Save individual HTML
        fig_net.write_html(str(export_dir / f"{prefix_id}_Network_Alliances.html"))

        safe_div_id = f"net_{prefix_id}".replace(" ", "_").replace(".", "_")
        html_net = fig_net.to_html(full_html=False, include_plotlyjs=False, div_id=safe_div_id, default_height="100%",
                                   default_width="100%")

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
            local_members = [m for members in members_list if (m := members) in G.nodes]
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
            bar_colors = [cluster_color_map[f] for f in contribs_df['Faction']]

            fig_contribs = go.Figure(go.Bar(
                x=contribs_df['Contributions'].tolist(), y=contribs_df['Faction'].tolist(), orientation='h',
                marker=dict(color=bar_colors), hovertext=contribs_df['Members'].tolist(),
                hovertemplate="<b>%{y}</b><br>Contributions: %{x}<br><br><b>Members:</b><br>%{hovertext}<extra></extra>"
            ))
            fig_contribs.update_layout(title="Total TDoc Contributions per Faction", showlegend=False)

            fig_contribs.write_html(str(export_dir / f"{prefix_id}_Faction_Contributions.html"))
            html_cluster_contribs = fig_contribs.to_html(full_html=False, include_plotlyjs=False, default_height="100%",
                                                         default_width="100%")

            bubble_df = contribs_df[contribs_df['Contributions'] > 0].copy()
            if not bubble_df.empty:
                bubble_colors = [cluster_color_map[f] for f in bubble_df['Faction']]
                max_contrib = float(bubble_df['Contributions'].max())

                # ---> THE FIX: sizeref is strictly formatted, and data is parsed with .tolist() to dodge NumPy KeyErrors!
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

                fig_cohesion.write_html(str(export_dir / f"{prefix_id}_Faction_Cohesion.html"))
                html_cohesion_plot = fig_cohesion.to_html(full_html=False, include_plotlyjs=False,
                                                          default_height="100%", default_width="100%")

    return html_net, html_cluster_contribs, html_cohesion_plot, html_faction_list