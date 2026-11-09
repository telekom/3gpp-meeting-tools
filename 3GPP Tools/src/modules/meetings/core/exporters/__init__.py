"""Exporters package with lazy loading to prevent circular import loops."""
from typing import Any

__all__ = [
    "PostMeetingExporterThread",
    "PreMeetingExporterThread",
    "DossierExporterThread",
    "ExcelExporterThread",
    "MarkdownExporterThread",
    "TDocMarkdownExporterThread",
    "TDocsPackageExporterThread",
    "ExportManifest",
    "IncrementalPlan",
    "load_manifest",
    "save_manifest",
    "load_or_reconstruct_manifest",
    "plan_incremental_session",
    "merge_manifest_into_tree",
    "format_master_changelog_markdown",
    "format_local_wi_changelog_markdown",
    "is_export_complete",
    "sanitize_folder_name",
]

_EXPORTS = {
    "PostMeetingExporterThread": "modules.meetings.core.exporters.post_meeting_exporter",
    "PreMeetingExporterThread": "modules.meetings.core.exporters.pre_meeting_exporter",
    "DossierExporterThread": "modules.meetings.core.exporters.dossier_exporter",
    "ExcelExporterThread": "modules.meetings.core.exporters.excel_exporter",
    "MarkdownExporterThread": "modules.meetings.core.exporters.markdown_exporter",
    "TDocMarkdownExporterThread": "modules.meetings.core.exporters.llm_exporter",
    "TDocsPackageExporterThread": "modules.meetings.core.exporters.tdocs_package_exporter",
}


def __getattr__(name: str) -> Any:
    if name in _EXPORTS:
        import importlib
        mod = importlib.import_module(_EXPORTS[name])
        val = getattr(mod, name)
        globals()[name] = val
        return val
    if name in (
        "ExportManifest",
        "IncrementalPlan",
        "load_manifest",
        "save_manifest",
        "load_or_reconstruct_manifest",
        "plan_incremental_session",
        "merge_manifest_into_tree",
        "format_master_changelog_markdown",
        "format_local_wi_changelog_markdown",
        "is_export_complete",
        "sanitize_folder_name",
    ):
        from modules.meetings.core.exporters import export_manifest
        val = getattr(export_manifest, name)
        globals()[name] = val
        return val
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
