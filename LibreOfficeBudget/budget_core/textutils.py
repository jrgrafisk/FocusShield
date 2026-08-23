"""Small text helpers shared by the parsing/detection modules.

Everything here is pure stdlib so the package can run both inside
LibreOffice's bundled Python and as a normal command line script.
"""

from __future__ import annotations

import re
import unicodedata

__all__ = ["fold", "squeeze", "is_blank_row"]

# Characters that NFKD does not decompose but that we still want to see as
# their "plain" counterpart when matching keywords.
_LETTER_MAP = {
    "æ": "ae",
    "ø": "o",
    "ß": "ss",
    "đ": "d",
    "ł": "l",
    "þ": "th",
}

_WS_RE = re.compile(r"\s+")


def fold(text: object) -> str:
    """Normalise ``text`` for keyword comparisons.

    Case, accents and Nordic letters are removed so that ``"Café Øst"``,
    ``"CAFE OST"`` and ``"cafe øst"`` all compare equal.
    """
    if text is None:
        return ""
    s = str(text)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.casefold()
    for src, dst in _LETTER_MAP.items():
        if src in s:
            s = s.replace(src, dst)
    return _WS_RE.sub(" ", s).strip()


def squeeze(text: object) -> str:
    """Collapse whitespace but keep the original characters."""
    if text is None:
        return ""
    return _WS_RE.sub(" ", str(text)).strip()


def is_blank_row(row) -> bool:
    """True when every cell in ``row`` is empty or whitespace."""
    return not any(str(cell).strip() for cell in row)
