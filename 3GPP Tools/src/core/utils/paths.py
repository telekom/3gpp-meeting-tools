from pathlib import Path


# Canonical application paths. These are fixed for the lifetime of the process
# and therefore resolved once at import time without creating anything.
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
APP_DATA_ROOT = Path.home() / "3GPP_Delegate_Helper"
CACHE_ROOT = APP_DATA_ROOT / "cache"
SPECS_ROOT = APP_DATA_ROOT / "specs"
TEMP_ROOT = APP_DATA_ROOT / "temp"
DB_ROOT = PROJECT_ROOT / "db"
OUTPUT_ROOT = PROJECT_ROOT.parent / "output"
MEETINGS_ROOT = CACHE_ROOT


def get_project_root() -> Path:
    """Return the absolute root of the '3GPP Tools' application package."""
    return PROJECT_ROOT


def get_db_root() -> Path:
    """Return the canonical directory containing SQLite databases (`src/db`)."""
    return DB_ROOT


def get_output_root() -> Path:
    """Return the default export directory for converted diagrams and documents (`output/`)."""
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    return OUTPUT_ROOT


def get_meetings_root() -> Path:
    """Return the canonical local cache directory for meetings and TDocs."""
    MEETINGS_ROOT.mkdir(parents=True, exist_ok=True)
    return MEETINGS_ROOT


def get_app_data_root() -> Path:
    """Return the application-owned persistent data root."""
    return APP_DATA_ROOT


def get_cache_root() -> Path:
    """Return the canonical application cache directory."""
    return CACHE_ROOT


def get_specs_root() -> Path:
    """Return the canonical local specification storage directory."""
    return SPECS_ROOT


def get_temp_root() -> Path:
    """Return the canonical application temporary-files directory."""
    return TEMP_ROOT


def ensure_app_data_dirs() -> None:
    """Create the minimum application-owned directory structure if necessary."""
    for path in (APP_DATA_ROOT, CACHE_ROOT, SPECS_ROOT, TEMP_ROOT):
        path.mkdir(parents=True, exist_ok=True)


def to_extended_path(p) -> str:
    """
    Convert a path to Windows extended-length format (\\?\\ or \\?\\UNC\\)
    to bypass the Win32 MAX_PATH (260 character) limit.
    """
    import os
    p_str = os.path.abspath(str(p))
    if os.name == "nt" and not p_str.startswith("\\\\?\\"):
        if p_str.startswith("\\\\"):
            return "\\\\?\\UNC\\" + p_str[2:]
        return "\\\\?\\" + p_str
    return p_str



def strip_extended_path(p) -> str:
    """Remove Windows \\\\?\\ or \\\\?\\UNC\\\\ prefix for clean relative paths, display, and external tools."""
    s = str(p)
    if s.startswith("\\\\?\\UNC\\"):
        return "\\\\" + s[8:]
    if s.startswith("\\\\?\\"):
        return s[4:]
    return s
