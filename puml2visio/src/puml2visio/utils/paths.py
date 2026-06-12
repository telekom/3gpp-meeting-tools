from pathlib import Path

def get_project_root() -> Path:
    """
    Safely resolves the absolute path to the root 'puml2visio' package,
    no matter where the app is launched from.
    """
    # __file__ is src/puml2visio/utils/paths.py
    # .parent is utils/
    # .parent.parent is puml2visio/ (The package root)
    return Path(__file__).resolve().parent.parent

def get_asset_path(filename: str) -> Path:
    """Returns the path to static assets like the JAR or templates."""
    # Example: Resolves to src/puml2visio/templates/plantuml.jar
    return get_project_root() / "assets" / filename