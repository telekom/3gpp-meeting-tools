# --- File: src/core/config/plot_styles.py ---
"""Plot styling constants and semantic color palettes for 3GPP analytics."""
import plotly.express as px

# Application unified theme colors
THEME_COLOR = '#1E5C99'
PALETTE = px.colors.qualitative.Plotly
CLUSTER_PALETTE = px.colors.qualitative.Alphabet

# Semantic color palette for standard 3GPP contribution outcomes
STATUS_COLORS = {
    'Agreed': '#10B981',       # Emerald Green
    'Approved': '#059669',     # Deep Green
    'Not Treated': '#94A3B8',  # Slate Grey
    'Revised': '#6366F1',      # Indigo Blue
    'Noted': '#F59E0B',        # Amber
    'Postponed': '#EAB308',    # Yellow
    'Merged': '#06B6D4',       # Cyan
    'Endorsed': '#14B8A6',     # Teal
    'Replied To': '#8B5CF6',   # Purple
    'Available': '#38BDF8',    # Sky Blue
    'Withdrawn': '#EF4444',    # Red
    'Unknown': '#CBD5E1',      # Muted Grey
}

STATUS_ORDER = [
    'Agreed', 'Approved', 'Merged', 'Endorsed',
    'Noted', 'Postponed', 'Replied To', 'Available',
    'Not Treated', 'Revised', 'Withdrawn', 'Unknown'
]

# Standard 3GPP Working Group color palette
WG_COLORS = {
    'SA1': '#3B82F6',    # Blue
    'SA2': '#10B981',    # Green
    'SA3': '#8B5CF6',    # Purple
    'SA4': '#F59E0B',    # Amber
    'SA5': '#06B6D4',    # Cyan
    'SA6': '#EC4899',    # Pink
    'RAN1': '#F97316',   # Orange
    'RAN2': '#6366F1',   # Indigo
    'RAN3': '#14B8A6',   # Teal
    'RAN4': '#84CC16',   # Lime
    'CT1': '#E11D48',    # Rose
    'CT3': '#A855F7',    # Violet
    'CT4': '#0284C7',    # Sky
}

# Document category colors
DOC_TYPE_COLORS = {
    'CR': '#1E5C99',          # Primary Blue
    'pCR': '#0284C7',         # Sky Blue
    'Discussion': '#F59E0B',  # Amber
    'LS in': '#10B981',       # Emerald Green
    'LS out': '#059669',      # Deep Green
    'LS': '#14B8A6',          # Teal
    'WID/SID': '#8B5CF6',     # Purple
    'Report': '#64748B',      # Slate
    'Specification': '#EC4899', # Pink
    'Other': '#94A3B8',       # Muted Grey
}

FALLBACK_PALETTE = ['#0D9488', '#2563EB', '#D97706', '#7C3AED', '#DB2777', '#059669', '#475569']


def normalize_status(val: str) -> str:
    """Normalizes raw 3GPP outcome strings into canonical categories."""
    s = str(val or '').strip().lower()
    if not s or s in ['nan', 'none', '-', '']:
        return 'Unknown'
    if 'not treated' in s or 'not handled' in s or 'left over' in s:
        return 'Not Treated'
    if 'agreed' in s:
        return 'Agreed'
    if 'approved' in s:
        return 'Approved'
    if 'revised' in s or 'revision' in s:
        return 'Revised'
    if 'noted' in s:
        return 'Noted'
    if 'postponed' in s:
        return 'Postponed'
    if 'merged' in s:
        return 'Merged'
    if 'endorsed' in s:
        return 'Endorsed'
    if 'replied' in s:
        return 'Replied To'
    if 'available' in s:
        return 'Available'
    if 'withdrawn' in s:
        return 'Withdrawn'
    return s.title()


def classify_doc_type(val: str, spec: str = "") -> str:
    """Classifies a document into a standard 3GPP document category."""
    s = str(val or '').strip().lower()
    if not s or s in ['none', 'nan', '-']:
        if spec and str(spec).strip():
            return 'CR'
        return 'Other'
    if 'pcr' in s or 'p-cr' in s or 'draftcr' in s or 'draft cr' in s:
        return 'pCR'
    if 'cr' in s:
        return 'CR'
    if 'ls in' in s or 'incoming ls' in s:
        return 'LS in'
    if 'ls out' in s or 'outgoing ls' in s:
        return 'LS out'
    if s == 'ls' or 'liaison' in s:
        return 'LS'
    if 'disc' in s or 'pdp' in s or 'decision' in s or 'proposal' in s or 'input' in s:
        return 'Discussion'
    if 'wid' in s or 'sid' in s or 'work item' in s or 'work plan' in s or 'study item' in s:
        return 'WID/SID'
    if 'report' in s:
        return 'Report'
    if 'spec' in s or 'ts' in s or 'tr' in s:
        return 'Specification'
    return s.title()
