from pathlib import Path


# Canonical application paths. These are fixed for the lifetime of the process
# and therefore resolved once at import time without creating anything.
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
APP_DATA_ROOT = Path.home() / "3GPP_Delegate_Helper"
CACHE_ROOT = APP_DATA_ROOT / "cache"
SPECS_ROOT = APP_DATA_ROOT / "specs"
TEMP_ROOT = APP_DATA_ROOT / "temp"


def get_project_root() -> Path:
    """Return the absolute root of the '3GPP Tools' application package."""
    return PROJECT_ROOT


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
