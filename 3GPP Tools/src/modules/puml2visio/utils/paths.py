from pathlib import Path

def get_project_root() -> Path:
    """
    Safely resolves the absolute path to the root '3GPP Tools' package,
    no matter where the app is launched from.
    """
    # __file__ is src/3GPP Tools/utils/paths.py
    # .parent is utils/
    # .parent.parent is 3GPP Tools/ (The package root)
    return Path(__file__).resolve().parent.parent.parent

def get_asset_path(filename: str) -> Path:
    """Returns the path to static assets like the JAR or templates."""
    return Path(__file__).resolve().parent.parent / "assets" / filename
