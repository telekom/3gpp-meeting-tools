import os
import logging
from pathlib import Path

def file_version_to_version(file_version: str) -> str:
    """
    Converts the file version of a 3GPP spec, e.g., a00 to a version number, e.g., 10.0.0.
    Args:
        file_version: A three-letter string containing the file version.

    Returns: The specification version number
    """

    def letter_to_number(character: str) -> str:
        if character.isdigit():
            number = character
        else:
            number = '{0}'.format(ord(character.lower()) - ord('a') + 10)
        return number

    if not file_version or len(file_version) < 3:
        return ""

    try:
        major_version = letter_to_number(file_version[0])
        middle_version = letter_to_number(file_version[1])
        minor_version = letter_to_number(file_version[2])
        return '{0}.{1}.{2}'.format(major_version, middle_version, minor_version)
    except Exception:
        return ""


def normalize_dynareport_spec_number(spec_number: str) -> str:
    """Return the canonical 3GPP DynaReport identifier for a specification.

    Examples:
        23.501 -> 23501
        23.801-01 -> 23801-01
        TR 23.801-01 -> 23801-01
    """
    import re

    clean_spec = str(spec_number or "").strip()
    clean_spec = re.sub(r'^(?:3GPP\s+)?(?:TS|TR)\s*', '', clean_spec, flags=re.IGNORECASE).strip()
    match = re.fullmatch(r'(\d{2})\.?(\d{3})(?:[-_.](\d{1,2}))?', clean_spec)
    if not match:
        return ""

    series, core, part = match.groups()
    identifier = f"{series}{core}"
    if part:
        identifier += f"-{part}"
    return identifier


def build_dynareport_url(spec_number: str) -> str:
    """Build the canonical DynaReport URL while preserving a specification part."""
    identifier = normalize_dynareport_spec_number(spec_number)
    return f"https://www.3gpp.org/DynaReport/{identifier}.htm" if identifier else ""

def open_extracted_documents(directory_path: Path) -> list:
    """
    Scans a directory for valid 3GPP document files and opens them natively.
    Actively ignores macOS hidden files and folder artifacts commonly found in ZIPs.
    """
    opened_files = []
    if not directory_path.exists() or not directory_path.is_dir():
        return opened_files

    # Strictly define which files we are allowed to automatically open
    allowed_extensions = {'.doc', '.docx', '.yaml', '.yml'}

    # rglob('*') recursively searches even if the zip extracted into a subfolder
    for file_path in directory_path.rglob('*'):
        if not file_path.is_file():
            continue

        # Ignore macOS hidden files and indexing artifacts
        if '__MACOSX' in file_path.parts or file_path.name.startswith('._'):
            continue

        # Check if the file is an allowed type
        if file_path.suffix.lower() in allowed_extensions:
            try:
                # Tell Windows to open the file using its default program (Word, Notepad++, etc.)
                os.startfile(str(file_path))
                opened_files.append(file_path)
            except Exception as e:
                logging.error(f"Failed to open {file_path}: {e}")

    return opened_files