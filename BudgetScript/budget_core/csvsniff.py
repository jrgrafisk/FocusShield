"""Turn "some CSV file" into a clean table.

Handles the parts of real world exports that break naive ``csv.DictReader``
usage: byte order marks, UTF-16, Latin-1, semicolon/comma/tab/pipe
separators, ``sep=;`` hint lines, bank preambles above the real header,
missing headers and ragged rows.
"""

from __future__ import annotations

import codecs
import csv
import io
import os
from collections import Counter
from typing import List, Optional, Sequence, Tuple

from .parsing import looks_like_amount, looks_like_date
from .textutils import fold, is_blank_row

__all__ = ["Table", "read_table", "detect_encoding", "detect_delimiter"]

CANDIDATE_DELIMITERS = (";", ",", "\t", "|")

# Words that give away a header row.  Compared token by token on folded text.
HEADER_WORDS = {
    "dato", "date", "datum", "bogfort", "bogforingsdato", "posteringsdato",
    "transaktionsdato", "valor", "valordato", "rentedato", "tid", "time",
    "tekst", "text", "beskrivelse", "description", "posteringstekst",
    "narrative", "details", "detaljer", "modtager", "afsender", "payee",
    "merchant", "reference", "memo", "note", "navn", "name", "titel",
    "belob", "amount", "sum", "value", "betrag", "montant", "beloeb",
    "saldo", "balance", "beholdning", "konto", "account", "kontonummer",
    "valuta", "currency", "type", "kategori", "category", "status",
    "debet", "kredit", "debit", "credit", "ind", "ud", "indsat", "haevet",
    "indbetaling", "udbetaling", "withdrawal", "deposit", "posting",
    "transaktion", "transaction", "art", "nummer", "number", "id",
    "afstemt", "gebyr", "fee",
}

# Sniffing never looks at more than this many rows.
SAMPLE_ROWS = 200


class Table(object):
    """A rectangular table of strings plus what we learned while reading it."""

    def __init__(self, rows, header=None, encoding="utf-8", delimiter=";",
                 quotechar='"', preamble=None, notes=None, source=""):
        self.rows: List[List[str]] = rows
        self.header: Optional[List[str]] = header
        self.encoding = encoding
        self.delimiter = delimiter
        self.quotechar = quotechar
        self.preamble: List[List[str]] = preamble or []
        self.notes: List[str] = notes or []
        self.source = source
        self.n_columns = max([len(r) for r in rows] + [len(header or [])] + [0])

    # -- convenience ------------------------------------------------------
    def column(self, index: int, limit: Optional[int] = None) -> List[str]:
        """All values of one column (as strings)."""
        rows = self.rows if limit is None else self.rows[:limit]
        return [r[index] if index < len(r) else "" for r in rows]

    def header_name(self, index: int) -> str:
        if self.header and index < len(self.header):
            return self.header[index].strip()
        return "Kolonne %d" % (index + 1)

    def header_names(self) -> List[str]:
        return [self.header_name(i) for i in range(self.n_columns)]

    def sample(self, count: int = 5) -> List[List[str]]:
        return self.rows[:count]

    def __len__(self) -> int:
        return len(self.rows)

    def describe(self) -> str:
        return "%d rækker, %d kolonner, tegnsæt %s, skilletegn %r" % (
            len(self.rows), self.n_columns, self.encoding,
            "TAB" if self.delimiter == "\t" else self.delimiter)

    # -- construction -----------------------------------------------------
    @classmethod
    def from_rows(cls, raw_rows: Sequence[Sequence[object]], has_header=None,
                  **kwargs) -> "Table":
        """Build a table from already split rows (CSV rows or sheet cells)."""
        rows = [[("" if c is None else str(c)) for c in row] for row in raw_rows]
        notes: List[str] = list(kwargs.pop("notes", []) or [])
        start, n_columns = _find_start(rows)
        preamble = [r for r in rows[:start] if not is_blank_row(r)]
        if preamble:
            notes.append("Sprang %d indledende linje(r) over." % len(preamble))

        body = [r for r in rows[start:] if not is_blank_row(r)]
        if not body:
            return cls([], None, preamble=preamble, notes=notes, **kwargs)

        header = None
        first = body[0]
        use_header = has_header
        if use_header is None:
            use_header = looks_like_header(first, body[1:21])
        if use_header:
            header = [str(c).strip() for c in first]
            body = body[1:]
            header = _uniquify(header)
        else:
            notes.append("Ingen overskriftsrække fundet - kolonnerne navngives automatisk.")

        ragged = 0
        fixed: List[List[str]] = []
        for row in body:
            if len(row) < n_columns:
                if len(row) < 2:
                    continue
                ragged += 1
                row = list(row) + [""] * (n_columns - len(row))
            fixed.append(list(row))
        if ragged:
            notes.append("%d række(r) havde færre kolonner end resten og blev fyldt ud." % ragged)

        return cls(fixed, header, preamble=preamble, notes=notes, **kwargs)


def _uniquify(names: Sequence[str]) -> List[str]:
    """Make header names unique and non empty."""
    seen: dict = {}
    out: List[str] = []
    for i, name in enumerate(names):
        name = (name or "").strip() or "Kolonne %d" % (i + 1)
        if name in seen:
            seen[name] += 1
            name = "%s (%d)" % (name, seen[name])
        else:
            seen[name] = 1
        out.append(name)
    return out


# --------------------------------------------------------------------------
# Encoding
# --------------------------------------------------------------------------

_BOMS = (
    (codecs.BOM_UTF8, "utf-8-sig"),
    (codecs.BOM_UTF32_LE, "utf-32"),
    (codecs.BOM_UTF32_BE, "utf-32"),
    (codecs.BOM_UTF16_LE, "utf-16"),
    (codecs.BOM_UTF16_BE, "utf-16"),
)


def detect_encoding(data: bytes, encoding: Optional[str] = None) -> Tuple[str, str]:
    """Decode ``data``; return ``(text, encoding_name)``.

    Never raises - the last resort (``cp1252``/``latin-1``) always decodes.
    """
    if encoding:
        return data.decode(encoding, "replace"), encoding

    for bom, name in _BOMS:
        if data.startswith(bom):
            try:
                return data.decode(name), name
            except UnicodeDecodeError:
                pass

    head = data[:4096]
    if head.count(b"\x00") > len(head) // 4:
        # UTF-16 without BOM: even NUL positions -> big endian.
        even = sum(1 for i in range(0, len(head) - 1, 2) if head[i] == 0)
        name = "utf-16-be" if even > len(head) // 5 else "utf-16-le"
        try:
            return data.decode(name), name
        except UnicodeDecodeError:
            pass

    for name in ("utf-8", "cp1252", "iso-8859-1"):
        try:
            return data.decode(name), name
        except UnicodeDecodeError:
            continue
    return data.decode("iso-8859-1", "replace"), "iso-8859-1"


# --------------------------------------------------------------------------
# Delimiter
# --------------------------------------------------------------------------

def detect_delimiter(text: str,
                     candidates: Sequence[str] = CANDIDATE_DELIMITERS) -> str:
    """Pick the separator that yields the most consistent, widest table."""
    lines = [ln for ln in text.splitlines() if ln.strip()][:60]
    if not lines:
        return ";"

    best = None
    best_score = 0.0
    for delimiter in candidates:
        try:
            rows = list(csv.reader(lines, delimiter=delimiter, quotechar='"'))
        except csv.Error:
            continue
        widths = [len(r) for r in rows if any(str(c).strip() for c in r)]
        if not widths:
            continue
        counts = Counter(widths)
        width, freq = max(counts.items(), key=lambda kv: (kv[1], kv[0]))
        if width < 2:
            continue
        score = (freq / float(len(widths))) * min(width, 12)
        if score > best_score + 1e-9:
            best_score = score
            best = delimiter
    return best or ";"


_SEP_HINT_PREFIXES = ("sep=", "SEP=")


def _strip_sep_hint(text: str) -> Tuple[str, Optional[str]]:
    """Consume Excel's ``sep=;`` first line if present."""
    first, sep, rest = text.partition("\n")
    if not sep:
        return text, None
    stripped = first.strip().lstrip("﻿")
    for prefix in _SEP_HINT_PREFIXES:
        if stripped.startswith(prefix) and len(stripped) == len(prefix) + 1:
            return rest, stripped[len(prefix)]
    return text, None


# --------------------------------------------------------------------------
# Header detection
# --------------------------------------------------------------------------

def looks_like_header(row: Sequence[str], following: Sequence[Sequence[str]]) -> bool:
    """Decide whether ``row`` is a header rather than the first data row."""
    cells = [str(c).strip() for c in row]
    non_empty = [c for c in cells if c]
    if not non_empty:
        return False

    if any(looks_like_date(c) for c in cells):
        return False

    hits = 0
    for cell in non_empty:
        folded = fold(cell)
        if folded in HEADER_WORDS:
            hits += 1
            continue
        tokens = [t for t in folded.replace("/", " ").replace("-", " ").split() if t]
        if any(t in HEADER_WORDS for t in tokens):
            hits += 1
    if hits >= 2 or (hits == 1 and len(non_empty) <= 3):
        return True

    following = [r for r in following if not is_blank_row(r)]
    if following:
        with_date = sum(1 for r in following if any(looks_like_date(c) for c in r))
        if with_date >= max(1, len(following) // 2):
            return True  # rows below have dates, this one does not
        numeric_below = sum(
            1 for r in following if any(looks_like_amount(c) for c in r))
        if numeric_below >= max(1, len(following) // 2) and \
                not any(looks_like_amount(c) for c in cells):
            return True
    return hits > 0


def _find_start(rows: Sequence[Sequence[str]]) -> Tuple[int, int]:
    """Find where the real table starts and how wide it is."""
    widths = Counter(len(r) for r in rows if not is_blank_row(r))
    if not widths:
        return 0, 0
    width = max(widths.items(), key=lambda kv: (kv[1], kv[0]))[0]

    for i, row in enumerate(rows):
        if is_blank_row(row) or len(row) != width:
            continue
        following = [r for r in rows[i + 1:i + 5] if not is_blank_row(r)]
        if not following or sum(1 for r in following if len(r) == width) >= \
                max(1, len(following) - 1):
            return i, width
    return 0, width


# --------------------------------------------------------------------------
# Public entry point
# --------------------------------------------------------------------------

def read_table(path: Optional[str] = None, data: Optional[bytes] = None,
               encoding: Optional[str] = None, delimiter: Optional[str] = None,
               has_header: Optional[bool] = None) -> Table:
    """Read ``path`` (or raw ``data``) and return a :class:`Table`.

    All detection can be overridden through the keyword arguments.
    """
    if data is None:
        if not path:
            raise ValueError("read_table kræver enten path eller data")
        with open(path, "rb") as fp:
            data = fp.read()

    text, encoding_name = detect_encoding(data, encoding)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text, hinted = _strip_sep_hint(text)
    if hinted and not delimiter:
        delimiter = hinted
    if not delimiter:
        delimiter = detect_delimiter(text)

    try:
        raw_rows = list(csv.reader(io.StringIO(text, newline=""),
                                   delimiter=delimiter, quotechar='"',
                                   skipinitialspace=True))
    except csv.Error as exc:  # pragma: no cover - only for pathological files
        raise ValueError("Kunne ikke læse CSV-filen: %s" % exc)

    table = Table.from_rows(raw_rows, has_header=has_header,
                            encoding=encoding_name, delimiter=delimiter,
                            source=os.path.basename(path) if path else "")
    return table
