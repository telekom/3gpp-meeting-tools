from collections import defaultdict
from typing import Dict, List


def invert_dict_defaultdict(input_dict: Dict[str, List[str]]) -> Dict[str, List[str]]:
    """
    Inverts a dictionary
    Args:
        input_dict: Dictionary with key=string, value=list of string

    Returns: An inverse dictionary

    """
    inverted_dict: Dict[str, List[str]] = defaultdict(list)
    for key, value_list in input_dict.items():
        for value in value_list:
            inverted_dict[value].append(key)
    return dict(inverted_dict)  # Convert back to a regular dict if needed
