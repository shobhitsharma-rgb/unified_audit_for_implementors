"""Turn a pasted list of IDs into a clean list.

People paste IDs typed with commas, or copied straight out of an Excel column
(one per line), or a mixture of both. All of these separate IDs:
comma, semicolon, newline, tab.

A SPACE does not: some IDs carry one ("123 456"), and the Selective Employee
Extractor matches those deliberately. Spaces around an ID are trimmed.
"""
import re

_SEPARATORS = re.compile(r"[,;\t\r\n]+")


def split_ids(text):
    """IDs in the order pasted, each once, blanks dropped.

    >>> split_ids("1020, BH0KS5HPZ\\n8OSU7337G;1020")
    ['1020', 'BH0KS5HPZ', '8OSU7337G']
    """
    seen, out = set(), []
    for part in _SEPARATORS.split(text or ""):
        part = part.strip()
        if part and part not in seen:
            seen.add(part)
            out.append(part)
    return out
