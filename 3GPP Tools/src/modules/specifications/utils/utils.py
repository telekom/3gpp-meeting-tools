# --- Add to: core/utils/utils.py ---

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