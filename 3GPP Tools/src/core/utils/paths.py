from pathlib import Path


def get_project_root() -> Path:
    """Return the absolute root of the 3GPP Tools source package."""
    return Path(__file__).resolve().parent.parent.parent


def get_app_data_root() -> Path:
    """Return the canonical per-user data root for 3GPP Delegate Helper."""
    return Path.home() / "3GPP_Delegate_Helper"


def get_cache_root() -> Path:
    """Return the default application cache directory."""
    return get_app_data_root() / "cache"


def get_specs_root() -> Path:
    """Return the default persistent specification storage directory."""
    return get_app_data_root() / "specs"


def get_temp_root() -> Path:
    """Return the application-owned temporary-work directory, creating it if needed."""
    path = get_app_data_root() / "temp"
    path.mkdir(parents=True, exist_ok=True)
    return path
