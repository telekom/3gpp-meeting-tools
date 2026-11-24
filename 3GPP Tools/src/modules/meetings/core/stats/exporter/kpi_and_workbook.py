"""KPI calculation, dataframe preparation, and Excel workbook export for meeting statistics."""

from __future__ import annotations

import re
from pathlib import Path
from typing import List, Tuple

import pandas as pd

from core.config.plot_styles import STATUS_ORDER, classify_doc_type, normalize_status
from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.core.agenda_manager import AgendaManager
from modules.meetings.core.stats.plot_insights import (
    _is_revision,
    build_company_specialization,
    build_revision_stats,
)


def compute_scope_kpis(scope_df: pd.DataFrame, scope_type: str = "global") -> dict:
    """Compute summary KPI cards and scoped dimension metadata for a view."""
    total_tdocs = len(scope_df)
    if total_tdocs == 0:
        return {
            "total_tdocs": 0,
            "orig_count": 0,
            "rev_count": 0,
            "cr_total": 0,
            "cr_agreed_count": 0,
            "cr_agree_pct": None,
            "handling_pct": 0.0,
            "handled_count": 0,
            "untreated_count": 0,
            "disc_total": 0,
            "ls_total": 0,
            "total_companies": 0,
            "joint_pct": 0.0,
            "scope_dim_val": "0",
            "scope_dim_title": "Items",
            "scope_dim_subtext": "",
        }

    is_rev = scope_df["_IsRevision"] if "_IsRevision" in scope_df.columns else scope_df.apply(_is_revision, axis=1)
    rev_count = int(is_rev.sum())
    orig_count = total_tdocs - rev_count

    doc_types = scope_df["Clean_Doc_Type"] if "Clean_Doc_Type" in scope_df.columns else pd.Series(["Other"] * total_tdocs)
    statuses = scope_df["Clean_Status"] if "Clean_Status" in scope_df.columns else scope_df["TDoc Status"].apply(normalize_status)

    cr_mask = doc_types.isin(["CR", "pCR"])
    cr_total = int(cr_mask.sum())
    cr_agreed_count = int((cr_mask & statuses.isin(["Agreed", "Approved"])).sum())
    cr_agree_pct = round((cr_agreed_count / cr_total) * 100, 1) if cr_total > 0 else None

    disc_mask = doc_types == "Discussion"
    disc_total = int(disc_mask.sum())

    ls_mask = doc_types.isin(["LS in", "LS out", "LS"])
    ls_total = int(ls_mask.sum())

    handled_mask = ~statuses.isin(["Not Treated", "Postponed", "Unknown"])
    handled_count = int(handled_mask.sum())
    untreated_count = total_tdocs - handled_count
    handling_pct = round((handled_count / total_tdocs) * 100, 1) if total_tdocs > 0 else 0.0

    companies = [c for sublist in scope_df.get("Clean_Companies", []) for c in (sublist or [])]
    total_companies = len(set(companies))

    solo_count = sum(1 for comps in scope_df.get("Clean_Companies", []) if len(comps or []) <= 1)
    joint_count = total_tdocs - solo_count
    joint_pct = round((joint_count / total_tdocs) * 100, 1) if total_tdocs > 0 else 0.0

    unique_wis = len(
        set(
            wi
            for sublist in scope_df.get("Clean_WI_List", [])
            for wi in (sublist or [])
            if wi not in ["Unspecified", "DUMMY", "None", ""]
        )
    )
    unique_ais = (
        scope_df.get("Agenda Item", pd.Series(dtype="object"))
        .fillna("")
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )

    if scope_type == "global":
        scope_dim_val = str(unique_wis)
        scope_dim_title = "Active Work Items"
        scope_dim_subtext = f"{unique_ais} Agenda Items Covered"
    elif scope_type == "wi":
        scope_dim_val = str(unique_ais)
        scope_dim_title = "Agenda Items"
        scope_dim_subtext = "Covered under this Work Item"
    else:  # "ai"
        scope_dim_val = str(unique_wis)
        scope_dim_title = "Related Work Items"
        scope_dim_subtext = "Addressed under this Agenda Item"

    return {
        "total_tdocs": total_tdocs,
        "orig_count": orig_count,
        "rev_count": rev_count,
        "cr_total": cr_total,
        "cr_agreed_count": cr_agreed_count,
        "cr_agree_pct": cr_agree_pct,
        "handling_pct": handling_pct,
        "handled_count": handled_count,
        "untreated_count": untreated_count,
        "disc_total": disc_total,
        "ls_total": ls_total,
        "total_companies": total_companies,
        "joint_pct": joint_pct,
        "scope_dim_val": scope_dim_val,
        "scope_dim_title": scope_dim_title,
        "scope_dim_subtext": scope_dim_subtext,
        "unique_ais": unique_ais,
    }


def _extract_wi_list(value) -> List[str]:
    text = str(value or "").strip()
    if not text or text.lower() in {"none", "nan", "-", "unknown", "dummy"}:
        return []
    return [part.strip() for part in re.split(r"[,;\n]+", text) if part.strip()]


def prepare_statistics_dataframe(tdocs_data: list, meeting_dir: Path) -> Tuple[pd.DataFrame, List[dict]]:
    """Normalize TDoc records, enrich AI/company/WI columns, and build filter rows."""
    df = pd.DataFrame(tdocs_data)
    if df.empty:
        return df, []

    df = df[
        ~df["TDoc Status"].astype(str).str.lower().str.contains("withdrawn", na=False)
    ].copy().reset_index(drop=True)

    df["Clean_Status"] = df["TDoc Status"].apply(normalize_status)
    df["Clean_Doc_Type"] = df.apply(lambda r: classify_doc_type(r.get("Type", ""), r.get("Spec", "")), axis=1)
    df["_IsRevision"] = df.apply(_is_revision, axis=1)
    df["Clean_Companies"] = df["Source"].apply(CompanySanitizer.get_matching_contributors)

    # Optional semantic AI metadata from the meeting's local agenda.csv.
    # The raw Agenda Item remains the canonical grouping/filter key.
    agenda_map = AgendaManager.load_local_agenda(meeting_dir / "Agenda")
    ai_acronym_counts: dict[str, int] = {}
    for ai in df["Agenda Item"].fillna("").astype(str).str.strip().unique():
        item = agenda_map.get(ai)
        if item and item.acronym:
            ai_acronym_counts[item.acronym] = ai_acronym_counts.get(item.acronym, 0) + 1

    def ai_meta(ai_value):
        ai = str(ai_value or "").strip()
        item = agenda_map.get(ai)
        if not item:
            return pd.Series(
                {"AI_Acronym": "", "AI_Topic": "", "AI_Full_Description": "", "AI_Display": ai, "AI_Axis": ai}
            )
        acronym, topic = item.acronym or "", item.topic_suffix or ""
        if acronym and ai_acronym_counts.get(acronym, 0) > 1 and topic:
            display = f"{ai} ({acronym} – {topic})"
        elif acronym:
            display = f"{ai} ({acronym})"
        else:
            clean_desc = re.sub(r"\s+", " ", item.description).strip() if item else ""
            if clean_desc:
                short_desc = clean_desc if len(clean_desc) <= 30 else clean_desc[:27].rstrip() + "..."
                display = f"{ai} - {short_desc}"
            else:
                display = ai
        # Keep plot axes compact. Rich semantic metadata remains in hover/data tables.
        if acronym and ai_acronym_counts.get(acronym, 0) > 1 and topic:
            short_topic = topic if len(topic) <= 24 else topic[:21].rstrip() + "..."
            axis = f"{ai}<br>{short_topic}"
        elif acronym:
            axis = f"{ai}<br>{acronym}"
        else:
            clean_desc = re.sub(r"\s+", " ", item.description).strip() if item else ""
            if clean_desc:
                short_desc = clean_desc if len(clean_desc) <= 30 else clean_desc[:27].rstrip() + "..."
                axis = f"{ai}<br>{short_desc[:20].rstrip()}..." if len(short_desc) > 20 else f"{ai}<br>{short_desc}"
            else:
                axis = ai
        return pd.Series(
            {
                "AI_Acronym": acronym,
                "AI_Topic": topic,
                "AI_Full_Description": item.description,
                "AI_Display": display,
                "AI_Axis": axis,
            }
        )

    ai_meta_df = df["Agenda Item"].apply(ai_meta)
    for col in ai_meta_df.columns:
        df[col] = ai_meta_df[col]

    if "Related WIs" in df.columns:
        df["Clean_WI_List"] = df["Related WIs"].apply(_extract_wi_list)
    else:
        df["Clean_WI_List"] = [[] for _ in range(len(df))]

    # Compact browser-side dataset used for interactive cross-filtering.
    # Raw TDoc text is intentionally not embedded; only dimensions needed by the dashboard.
    interactive_rows: List[dict] = []
    for _, row in df.iterrows():
        interactive_rows.append(
            {
                "tdoc": str(row.get("TDoc", "") or ""),
                "type": str(row.get("Clean_Doc_Type", row.get("Type", "")) or ""),
                "ai": str(row.get("Agenda Item", "") or "").strip(),
                "aiDisplay": str(row.get("AI_Display", row.get("Agenda Item", "")) or "").strip(),
                "wis": list(row.get("Clean_WI_List", []) or []),
                "companies": list(row.get("Clean_Companies", []) or []),
                "source": str(row.get("Source", "") or ""),
                "status": str(row.get("TDoc Status", "") or "").strip(),
                "isRevision": bool(str(row.get("Is revision of", "") or "").strip())
                or bool(re.search(r"(?:r|rev)\d{1,2}[a-zA-Z]?$", str(row.get("TDoc", "") or ""), re.IGNORECASE)),
            }
        )
    return df, interactive_rows


def export_statistics_workbook(df: pd.DataFrame, path: Path) -> None:
    """Write reusable aggregated statistics behind the HTML dashboard."""
    kpis = compute_scope_kpis(df, "global")
    summary_data = [
        {"Metric": "Total Contributions (TDocs)", "Value": kpis["total_tdocs"], "Notes": "Excludes withdrawn documents"},
        {"Metric": "Original TDocs", "Value": kpis["orig_count"], "Notes": "First document revisions"},
        {"Metric": "Subsequent Revisions", "Value": kpis["rev_count"], "Notes": "r01, r02, etc."},
        {
            "Metric": "CR Agreement Rate",
            "Value": f"{kpis['cr_agree_pct']}%" if kpis["cr_agree_pct"] is not None else "N/A",
            "Notes": f"{kpis['cr_agreed_count']} of {kpis['cr_total']} CRs/pCRs agreed",
        },
        {
            "Metric": "Meeting Treatment Rate",
            "Value": f"{kpis['handling_pct']}%",
            "Notes": f"{kpis['handled_count']} handled, {kpis['untreated_count']} not treated/postponed",
        },
        {"Metric": "Change Requests (CR + pCR)", "Value": kpis["cr_total"], "Notes": "Normative specification changes"},
        {"Metric": "Discussion Papers", "Value": kpis["disc_total"], "Notes": "Inputs, PDPs, decisions"},
        {"Metric": "Liaison Statements (LS)", "Value": kpis["ls_total"], "Notes": "Incoming & outgoing LS"},
        {"Metric": "Active Contributing Entities", "Value": kpis["total_companies"], "Notes": "Sanitized company names"},
        {"Metric": "Joint Contributions Share", "Value": f"{kpis['joint_pct']}%", "Notes": "Multi-company co-sponsored docs"},
        {"Metric": "Active Work Items", "Value": kpis["scope_dim_val"], "Notes": "Tracked and general WIs"},
    ]
    summary_df = pd.DataFrame(summary_data)

    doc_type_summary = df.groupby(["Clean_Doc_Type", "Clean_Status"]).size().unstack(fill_value=0)
    ordered_st = [s for s in STATUS_ORDER if s in doc_type_summary.columns] + [
        c for c in doc_type_summary.columns if c not in STATUS_ORDER
    ]
    doc_type_summary = doc_type_summary[ordered_st]
    doc_type_summary["Total"] = doc_type_summary.sum(axis=1)
    doc_type_summary = (
        doc_type_summary.sort_values("Total", ascending=False).reset_index().rename(columns={"Clean_Doc_Type": "Document Type"})
    )

    ai_cols = ["Agenda Item", "AI_Acronym", "AI_Topic"]
    ai_volume = df.groupby(ai_cols, dropna=False).size().reset_index(name="TDocs").sort_values("TDocs", ascending=False)
    status = df["Clean_Status"].fillna("").astype(str).value_counts().rename_axis("Status").reset_index(name="TDocs")
    company_rows = []
    for _, row in df.iterrows():
        for company in set(row.get("Clean_Companies", []) or []):
            company_rows.append(
                {
                    "Company": company,
                    "Agenda Item": row.get("Agenda Item", ""),
                    "WI": "; ".join(row.get("Clean_WI_List", []) or []),
                    "TDoc": row.get("TDoc", ""),
                    "Type": row.get("Clean_Doc_Type", ""),
                    "Status": row.get("Clean_Status", ""),
                }
            )
    company_df = pd.DataFrame(company_rows)
    company_volume = (
        company_df.groupby("Company").size().reset_index(name="TDocs").sort_values("TDocs", ascending=False)
        if not company_df.empty
        else pd.DataFrame(columns=["Company", "TDocs"])
    )
    wi_rows = []
    for _, row in df.iterrows():
        for wi in row.get("Clean_WI_List", []) or []:
            wi_rows.append(
                {
                    "WI/SI": wi,
                    "TDoc": row.get("TDoc", ""),
                    "Agenda Item": row.get("Agenda Item", ""),
                    "Type": row.get("Clean_Doc_Type", ""),
                    "Status": row.get("Clean_Status", ""),
                }
            )
    wi_df = pd.DataFrame(wi_rows)
    wi_volume = (
        wi_df.groupby("WI/SI").size().reset_index(name="TDocs").sort_values("TDocs", ascending=False)
        if not wi_df.empty
        else pd.DataFrame(columns=["WI/SI", "TDocs"])
    )
    revision = build_revision_stats(df).drop(columns=["_Axis"], errors="ignore")
    specialization = build_company_specialization(df)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Executive Summary", index=False)
        doc_type_summary.to_excel(writer, sheet_name="Document Types", index=False)
        ai_volume.to_excel(writer, sheet_name="AI Volume", index=False)
        status.to_excel(writer, sheet_name="Outcomes", index=False)
        wi_volume.to_excel(writer, sheet_name="WI-SI Volume", index=False)
        company_volume.to_excel(writer, sheet_name="Company Volume", index=False)
        revision.to_excel(writer, sheet_name="Revision Activity", index=False)
        specialization.to_excel(writer, sheet_name="Specialization", index=False)
        if not company_df.empty:
            pd.crosstab(company_df["Company"], company_df["Agenda Item"]).to_excel(writer, sheet_name="Company x AI")
        if not wi_df.empty:
            pd.crosstab(wi_df["WI/SI"], wi_df["Agenda Item"]).to_excel(writer, sheet_name="WI-SI x AI")
