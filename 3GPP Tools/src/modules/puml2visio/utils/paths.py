from pathlib import Path


def get_puml2visio_asset_path(filename: str) -> Path:
    """Returns the path to static assets like the JAR or templates."""
    return Path(__file__).resolve().parent.parent / "assets" / filename
