"""
Utilities module.
"""
import re
from typing import Dict


def dict_replace(s: str, d: Dict) -> str:
    """
    Replaces all dictionary keys in a string with their respective dictionary values.
    
    :param s: The string
    :param d: The dictionary
    :return: The new string
    """
    pattern = re.compile('|'.join(d.keys()))
    return pattern.sub(lambda x: d[x.group()], s)
