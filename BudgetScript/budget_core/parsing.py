"""Tolerant parsing of the amounts and dates found in bank CSV exports.

The functions here are deliberately forgiving: real exports contain
currency symbols, non breaking spaces, trailing minus signs, parentheses
for negative numbers, four different thousand separators and about a
dozen date layouts.  Nothing raises - unparsable input returns ``None``
so the caller can count and report it.
"""

from __future__ import annotations

import datetime as _dt
import re
from typing import Iterable, List, Optional, Sequence, Tuple

from .textutils import fold

__all__ = [
    "parse_amount",
    "detect_decimal_separator",
    "looks_like_amount",
    "parse_date",
    "detect_dayfirst",
    "looks_like_date",
    "month_key",
    "MONTH_NAMES_DA",
]

# --------------------------------------------------------------------------
# Amounts
# --------------------------------------------------------------------------

# Currency markers that may sit before or after the number.
_CURRENCY_RE = re.compile(
    r"(?i)(?:^|(?<=[\s\d]))(?:dkk|sek|nok|isk|eur|usd|gbp|chf|pln|czk|huf|kr\.?|kroner)"
    r"(?=$|[\s\d.,+-])|[€$£¥₤]"
)
# Debit/credit markers used by some accounting systems ("1.234,56 DR").
_DC_SUFFIX_RE = re.compile(r"(?i)\s*(?:dr|db|cr|kr)\.?$")
_NUMBER_BODY_RE = re.compile(r"^[0-9]+(?:[.,'’][0-9]+)*$")
_GROUPED_RE_CACHE: dict = {}


def _clean_number_text(raw: object) -> Tuple[Optional[str], bool]:
    """Return ``(digits_and_separators, negative)`` for ``raw``.

    ``(None, False)`` is returned when ``raw`` does not look like a number
    at all.
    """
    if raw is None:
        return None, False
    if isinstance(raw, bool):
        return None, False
    if isinstance(raw, (int, float)):
        value = float(raw)
        return (repr(abs(value)), value < 0)

    s = str(raw).strip()
    if not s:
        return None, False

    # Normalise exotic spaces and dashes.
    for ch in (" ", " ", " ", " "):
        s = s.replace(ch, " ")
    for ch in ("−", "–", "—"):
        s = s.replace(ch, "-")

    s = _CURRENCY_RE.sub(" ", s)
    negative = False

    # "1.234,56 DR" / "1.234,56 CR"
    m = _DC_SUFFIX_RE.search(s)
    if m:
        if m.group(0).strip().lower().startswith(("dr", "db")):
            negative = True
        s = s[: m.start()]

    s = s.strip()
    if s.startswith("(") and s.endswith(")"):
        negative = not negative
        s = s[1:-1].strip()

    for prefix in ("-", "+"):
        if s.startswith(prefix):
            negative = negative != (prefix == "-")
            s = s[1:].strip()
            break
    for suffix in ("-", "+"):
        if s.endswith(suffix):
            negative = negative != (suffix == "-")
            s = s[:-1].strip()
            break

    s = s.replace(" ", "").replace("'", "").replace("’", "")
    if not s or not _NUMBER_BODY_RE.match(s):
        return None, False
    return s, negative


def _is_grouped(body: str, sep: str) -> bool:
    """True when ``sep`` is used as a thousand separator in ``body``."""
    pattern = _GROUPED_RE_CACHE.get(sep)
    if pattern is None:
        pattern = re.compile(r"^\d{1,3}(?:%s\d{3})+$" % re.escape(sep))
        _GROUPED_RE_CACHE[sep] = pattern
    return bool(pattern.match(body))


def _value_decimal_separator(body: str) -> Optional[str]:
    """Guess the decimal separator of a single cleaned number."""
    has_dot = "." in body
    has_comma = "," in body
    if has_dot and has_comma:
        return "." if body.rfind(".") > body.rfind(",") else ","
    for sep, other in ((",", "."), (".", ",")):
        if sep not in body:
            continue
        if body.count(sep) > 1:
            return other  # repeated separator can only be grouping
        tail = body.rsplit(sep, 1)[1]
        if len(tail) == 3 and _is_grouped(body, sep):
            return other  # "1.234" / "1,234" - most likely a thousand group
        return sep
    return None


def _assemble(body: str, decimal: str) -> str:
    grouping = "," if decimal == "." else "."
    body = body.replace(grouping, "")
    parts = body.split(decimal)
    if len(parts) == 1:
        return parts[0]
    if len(parts) > 2:
        # Separator used for grouping after all, except possibly the last one.
        if len(parts[-1]) in (1, 2):
            return "".join(parts[:-1]) + "." + parts[-1]
        return "".join(parts)
    return parts[0] + "." + parts[1]


def parse_amount(raw: object, decimal: Optional[str] = None) -> Optional[float]:
    """Parse ``raw`` into a float, or return ``None``.

    ``decimal`` forces the decimal separator (``","`` or ``"."``).  When it
    is omitted the separator is guessed from the value itself, which is
    reliable for values with both a grouping and a decimal separator but
    ambiguous for values such as ``"1,234"``.  Use
    :func:`detect_decimal_separator` on a whole column when possible.
    """
    body, negative = _clean_number_text(raw)
    if body is None:
        return None
    sep = decimal or _value_decimal_separator(body)
    if sep is None:
        text = body
    else:
        text = _assemble(body, sep)
    try:
        value = float(text)
    except ValueError:
        return None
    return -value if negative else value


def looks_like_amount(raw: object) -> bool:
    """True when ``raw`` can be read as a number."""
    return _clean_number_text(raw)[0] is not None


def detect_decimal_separator(samples: Iterable[object]) -> Optional[str]:
    """Decide whether a column uses ``,`` or ``.`` as decimal separator."""
    strong = {",": 0, ".": 0}
    weak = {",": 0, ".": 0}
    for raw in samples:
        body, _ = _clean_number_text(raw)
        if not body:
            continue
        has_dot = "." in body
        has_comma = "," in body
        if has_dot and has_comma:
            strong["." if body.rfind(".") > body.rfind(",") else ","] += 2
            continue
        for sep, other in ((",", "."), (".", ",")):
            if sep not in body:
                continue
            if body.count(sep) > 1:
                strong[other] += 1
            else:
                tail = body.rsplit(sep, 1)[1]
                if len(tail) in (1, 2) or len(tail) > 3:
                    strong[sep] += 1
                elif _is_grouped(body, sep):
                    weak[other] += 1
                else:
                    strong[sep] += 1
    if strong[","] != strong["."]:
        return "," if strong[","] > strong["."] else "."
    if weak[","] != weak["."]:
        return "," if weak[","] > weak["."] else "."
    if strong[","] or weak[","]:
        return ","
    return None


# --------------------------------------------------------------------------
# Dates
# --------------------------------------------------------------------------

MONTH_NAMES_DA = [
    "januar", "februar", "marts", "april", "maj", "juni",
    "juli", "august", "september", "oktober", "november", "december",
]

_MONTHS = {}
for _i, _name in enumerate(MONTH_NAMES_DA, start=1):
    _MONTHS[_name] = _i
    _MONTHS[_name[:3]] = _i
for _i, _name in enumerate(
    ["january", "february", "march", "april", "may", "june", "july",
     "august", "september", "october", "november", "december"], start=1
):
    _MONTHS[_name] = _i
    _MONTHS[_name[:3]] = _i
_MONTHS.update({"mar": 3, "okt": 10, "oct": 10, "sept": 9, "des": 12})

_TOKEN_RE = re.compile(r"[0-9]+|[a-zæøå]+", re.IGNORECASE)
_COMPACT_RE = re.compile(r"^(\d{4})(\d{2})(\d{2})$")
_COMPACT_DMY_RE = re.compile(r"^(\d{2})(\d{2})(\d{4})$")


def _make_date(year: int, month: int, day: int) -> Optional[_dt.date]:
    if year < 100:
        year += 2000 if year < 70 else 1900
    if not (1900 <= year <= 2200 and 1 <= month <= 12 and 1 <= day <= 31):
        return None
    try:
        return _dt.date(year, month, day)
    except ValueError:
        return None


def parse_date(raw: object, dayfirst: bool = True) -> Optional[_dt.date]:
    """Parse a date written in (almost) any common layout.

    Understood: ``2024-05-31``, ``31-05-2024``, ``31/05/24``, ``31.05.2024``,
    ``20240531``, ``31052024``, ``31. maj 2024``, ``May 31, 2024`` and any of
    those followed by a clock time.
    """
    if raw is None:
        return None
    if isinstance(raw, _dt.datetime):
        return raw.date()
    if isinstance(raw, _dt.date):
        return raw

    s = str(raw).strip()
    if not s:
        return None

    compact = s.replace(" ", "")
    m = _COMPACT_RE.match(compact)
    if m:
        result = _make_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if result:
            return result
    m = _COMPACT_DMY_RE.match(compact)
    if m:
        result = (_make_date(int(m.group(3)), int(m.group(2)), int(m.group(1))) or
                  _make_date(int(m.group(3)), int(m.group(1)), int(m.group(2))))
        if result:
            return result

    tokens = _TOKEN_RE.findall(s)
    if len(tokens) < 3:
        return None

    month_from_name = None
    numbers: List[int] = []
    for tok in tokens:
        if tok.isdigit():
            if len(tok) > 4:
                return None
            numbers.append(int(tok))
        else:
            name = fold(tok).rstrip(".")
            if name in _MONTHS and month_from_name is None:
                month_from_name = _MONTHS[name]
            elif numbers:
                break  # a word after the numbers - stop, e.g. "31-05-2024 kl"
        if month_from_name is not None and len(numbers) >= 2:
            break
        if month_from_name is None and len(numbers) >= 3:
            break

    if month_from_name is not None:
        if len(numbers) < 2:
            return None
        a, b = numbers[0], numbers[1]
        year, day = (a, b) if a > 31 else (b, a)
        return _make_date(year, month_from_name, day)

    if len(numbers) < 3:
        return None
    a, b, c = numbers[:3]
    if a > 31:  # year first
        return _make_date(a, b, c)
    if a > 12:
        return _make_date(c, b, a)
    if b > 12:
        return _make_date(c, a, b)
    if dayfirst:
        return _make_date(c, b, a) or _make_date(c, a, b)
    return _make_date(c, a, b) or _make_date(c, b, a)


def looks_like_date(raw: object) -> bool:
    """True when ``raw`` can be read as a date in some layout."""
    return parse_date(raw) is not None


def detect_dayfirst(samples: Iterable[object]) -> bool:
    """Decide between ``31/05/2024`` (day first) and ``05/31/2024``."""
    day_votes = 0
    month_votes = 0
    for raw in samples:
        if raw is None:
            continue
        s = str(raw).strip()
        if not s or _COMPACT_RE.match(s.replace(" ", "")):
            continue
        numbers = [int(t) for t in _TOKEN_RE.findall(s) if t.isdigit() and len(t) <= 4]
        if len(numbers) < 3:
            continue
        a, b, _c = numbers[:3]
        if a > 31:  # ISO style, tells us nothing
            continue
        if a > 12:
            day_votes += 1
        elif b > 12:
            month_votes += 1
    if month_votes > day_votes:
        return False
    return True


def month_key(date: _dt.date) -> str:
    """``datetime.date`` -> ``"2024-05"`` (sorts chronologically)."""
    return "%04d-%02d" % (date.year, date.month)


def month_label(key: str) -> str:
    """``"2024-05"`` -> ``"maj 2024"`` for human readable headings."""
    try:
        year, month = key.split("-")
        return "%s %s" % (MONTH_NAMES_DA[int(month) - 1], year)
    except (ValueError, IndexError):
        return key


def date_samples(values: Sequence[object], limit: int = 200) -> List[object]:
    """First ``limit`` non empty values - convenience for the detectors."""
    out: List[object] = []
    for value in values:
        if value is None:
            continue
        if str(value).strip():
            out.append(value)
        if len(out) >= limit:
            break
    return out
