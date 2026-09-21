from pathlib import Path

from core.utils.paths import get_project_root


def get_puml2visio_asset_path(filename: str) -> Path:
    """Returns the path to static assets like the JAR or templates."""
    return get_project_root() / "assets" / filename
