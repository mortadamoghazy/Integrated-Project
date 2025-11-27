"""
Helpers for normalizing labels and employee IDs.
"""

import re
import unicodedata


def _strip_accents(s: str) -> str:
    """Remove accents so labels can be compared reliably."""
    return "".join(
        c
        for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )


def _norm_label(s) -> str:
    """
    Clean and normalize header/field labels:
    - lowercase
    - remove accents
    - collapse whitespace
    - keep only simple characters
    """
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = _strip_accents(s)
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^\w\s./-]", "", s)
    return s


def _norm_emp_id(x, width: int) -> str:
    """
    Convert an employee ID into a zero-padded numeric string (e.g., 00014).
    """
    digits = re.sub(r"\D", "", str(x))
    return digits.zfill(width) if digits else ""
