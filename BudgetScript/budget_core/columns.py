"""Figure out which column holds the date, the text and the amount.

Bank exports differ wildly: one amount column with signs, two columns
(``Hævet``/``Indsat``), an extra ``Saldo`` column that must not be mistaken
for the amount, two date columns (``Bogført``/``Valør``), text spread over
a type column and a description column...  This module scores every column
and picks the most likely role, and it explains its choices so the user can
correct them in the import dialog.
"""

from __future__ import annotations

from typing import List, Optional, Sequence

from .parsing import (detect_dayfirst, detect_decimal_separator,
                      looks_like_amount, looks_like_date, parse_amount)
from .textutils import fold

__all__ = ["ColumnStats", "ColumnMapping", "analyse_columns", "detect_mapping"]

DATE_WORDS = ("dato", "date", "datum", "bogfort", "bogforingsdato", "posteringsdato",
              "transaktionsdato", "valor", "valordato", "posting", "booking", "fecha")
DATE_WORDS_PRIMARY = ("dato", "date", "bogfort", "posteringsdato", "transaktionsdato",
                      "posting", "booking")
TEXT_WORDS = ("tekst", "text", "beskrivelse", "description", "posteringstekst",
              "narrative", "details", "detaljer", "modtager", "afsender", "payee",
              "merchant", "reference", "memo", "note", "navn", "name", "titel",
              "forklaring", "kommentar", "bemaerkning")
AMOUNT_WORDS = ("belob", "beloeb", "amount", "sum", "betrag", "montant", "value",
                "transaktionsbelob", "bevaegelse", "posteringsbelob")
IN_WORDS = ("indsat", "indbetaling", "indbetalt", "ind", "kredit", "credit",
            "deposit", "modtaget", "tilgang", "indgaende", "income", "haevet ind")
OUT_WORDS = ("haevet", "udbetaling", "udbetalt", "ud", "debet", "debit",
             "withdrawal", "betalt", "afgang", "udgaende", "expense", "traek")
BALANCE_WORDS = ("saldo", "balance", "beholdning", "kontosaldo", "running balance",
                 "ny saldo")
CURRENCY_WORDS = ("valuta", "currency", "mont", "vaeluta")
IGNORE_WORDS = ("kontonummer", "account number", "kortnummer", "id", "reg nr",
                "registreringsnummer", "afstemt", "status")


def _has_word(header: str, words: Sequence[str]) -> bool:
    folded = fold(header)
    if not folded:
        return False
    if folded in words:
        return True
    tokens = [t for t in folded.replace("/", " ").replace("-", " ")
              .replace(".", " ").split() if t]
    if any(t in words for t in tokens):
        return True
    return any(w in folded for w in words if len(w) > 4)


class ColumnStats(object):
    """Simple per column statistics used by the detectors."""

    def __init__(self, index: int, header: str, values: Sequence[str]):
        self.index = index
        self.header = header or ""
        self.values = list(values)
        filled = [v for v in self.values if str(v).strip()]
        self.filled = filled
        total = float(len(filled)) or 1.0
        self.fill_ratio = len(filled) / float(len(self.values) or 1)
        self.date_ratio = sum(1 for v in filled if looks_like_date(v)) / total
        numeric = [v for v in filled if looks_like_amount(v)]
        self.numeric_ratio = len(numeric) / total
        self.decimal_ratio = (
            sum(1 for v in numeric if "," in str(v) or "." in str(v)) /
            (float(len(numeric)) or 1.0))
        self.avg_length = (sum(len(str(v)) for v in filled) / total) if filled else 0.0
        self.distinct_ratio = (len(set(str(v).strip().lower() for v in filled)) /
                               total) if filled else 0.0
        self.long_int_ratio = (
            sum(1 for v in numeric
                if str(v).strip().replace(" ", "").isdigit() and
                len(str(v).strip().replace(" ", "")) >= 8) /
            (float(len(numeric)) or 1.0))

    def is_dateish(self) -> bool:
        return self.date_ratio >= 0.6

    def is_numeric(self) -> bool:
        return self.numeric_ratio >= 0.6 and self.date_ratio < 0.6

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return "<ColumnStats %d %r date=%.2f num=%.2f>" % (
            self.index, self.header, self.date_ratio, self.numeric_ratio)


class ColumnMapping(object):
    """Which column plays which role, plus the parsing settings."""

    def __init__(self):
        self.date: Optional[int] = None
        self.text: List[int] = []
        self.amount: Optional[int] = None
        self.amount_in: Optional[int] = None
        self.amount_out: Optional[int] = None
        self.balance: Optional[int] = None
        self.currency: Optional[int] = None
        self.decimal: Optional[str] = None
        self.dayfirst: bool = True
        self.notes: List[str] = []

    @property
    def has_amount(self) -> bool:
        return self.amount is not None or self.amount_in is not None or \
            self.amount_out is not None

    def is_usable(self) -> bool:
        return self.date is not None and self.has_amount

    def describe(self, table) -> List[str]:
        """Human readable summary of the mapping (for the dialog/CLI)."""
        out = []
        if self.date is not None:
            out.append("Dato: %s" % table.header_name(self.date))
        if self.text:
            out.append("Tekst: %s" % " + ".join(table.header_name(i) for i in self.text))
        if self.amount is not None:
            out.append("Beløb: %s" % table.header_name(self.amount))
        if self.amount_in is not None:
            out.append("Indbetalt: %s" % table.header_name(self.amount_in))
        if self.amount_out is not None:
            out.append("Hævet: %s" % table.header_name(self.amount_out))
        if self.balance is not None:
            out.append("Saldo (ignoreres): %s" % table.header_name(self.balance))
        out.append("Decimaltegn: %s" % ("komma" if self.decimal == "," else "punktum"))
        out.append("Datoformat: %s" % ("dag før måned" if self.dayfirst else "måned før dag"))
        return out


def analyse_columns(table, limit: int = 300) -> List[ColumnStats]:
    """Compute :class:`ColumnStats` for every column of ``table``."""
    return [ColumnStats(i, table.header_name(i), table.column(i, limit))
            for i in range(table.n_columns)]


def _pick_date(stats: Sequence[ColumnStats]) -> Optional[int]:
    candidates = [s for s in stats if s.is_dateish()]
    if not candidates:
        candidates = [s for s in stats if s.date_ratio >= 0.3]
    if not candidates:
        return None
    named = [s for s in candidates if _has_word(s.header, DATE_WORDS_PRIMARY)]
    if named:
        return named[0].index
    named = [s for s in candidates if _has_word(s.header, DATE_WORDS)]
    if named:
        return named[0].index
    candidates.sort(key=lambda s: (-s.date_ratio, s.index))
    return candidates[0].index


def _looks_like_balance(amount_col: ColumnStats, balance_col: ColumnStats,
                        decimal: Optional[str]) -> bool:
    """True when ``balance_col`` behaves like a running balance of ``amount_col``."""
    hits = 0
    tested = 0
    prev = None
    for amount_raw, balance_raw in zip(amount_col.values, balance_col.values):
        balance = parse_amount(balance_raw, decimal)
        amount = parse_amount(amount_raw, decimal)
        if balance is None or amount is None:
            prev = balance if balance is not None else prev
            continue
        if prev is not None:
            tested += 1
            if abs((balance - prev) - amount) < 0.02 or \
                    abs((prev - balance) - amount) < 0.02:
                hits += 1
        prev = balance
    return tested >= 3 and hits >= tested * 0.7


def _pick_amount(stats: Sequence[ColumnStats], date_index: Optional[int],
                 mapping: ColumnMapping) -> None:
    numeric = [s for s in stats
               if s.is_numeric() and s.index != date_index and s.long_int_ratio < 0.8]
    if not numeric:
        mapping.notes.append("Fandt ingen talkolonne - vælg beløbskolonnen manuelt.")
        return

    # 1) An explicit balance column is never the amount.
    balances = [s for s in numeric if _has_word(s.header, BALANCE_WORDS)]
    if balances:
        mapping.balance = balances[0].index
        numeric = [s for s in numeric if s.index != mapping.balance]

    # 2) Separate in/out columns?
    ins = [s for s in numeric if _has_word(s.header, IN_WORDS)]
    outs = [s for s in numeric if _has_word(s.header, OUT_WORDS)]
    if ins and outs and ins[0].index != outs[0].index:
        mapping.amount_in = ins[0].index
        mapping.amount_out = outs[0].index
        return
    if len(numeric) == 2 and not any(_has_word(s.header, AMOUNT_WORDS) for s in numeric):
        first, second = numeric
        exclusive = 0
        tested = 0
        for a, b in zip(first.values, second.values):
            a_filled = bool(str(a).strip())
            b_filled = bool(str(b).strip())
            if a_filled or b_filled:
                tested += 1
                if a_filled != b_filled:
                    exclusive += 1
        if tested >= 3 and exclusive >= tested * 0.8:
            mapping.amount_in = first.index
            mapping.amount_out = second.index
            mapping.notes.append(
                "To kolonner udfyldes skiftevis - tolkes som ind- og udbetaling. "
                "Byt om i dialogen hvis det er omvendt.")
            return

    if len(numeric) == 1:
        mapping.amount = numeric[0].index
        return

    # 3) Named amount column wins.
    named = [s for s in numeric if _has_word(s.header, AMOUNT_WORDS)]
    if named:
        mapping.amount = named[0].index
        rest = [s for s in numeric if s.index != mapping.amount]
        if mapping.balance is None and rest:
            for other in rest:
                if _looks_like_balance(named[0], other, mapping.decimal):
                    mapping.balance = other.index
                    break
        return

    # 4) Otherwise look for the running balance pattern.
    for candidate in numeric:
        for other in numeric:
            if other.index == candidate.index:
                continue
            if _looks_like_balance(candidate, other, mapping.decimal):
                mapping.amount = candidate.index
                mapping.balance = other.index
                mapping.notes.append(
                    "Kolonnen \"%s\" ser ud til at være en løbende saldo og springes over."
                    % other.header)
                return

    # 5) Give up gracefully: prefer decimals, signs and a late position.
    def score(s: ColumnStats) -> tuple:
        values = [parse_amount(v, mapping.decimal) for v in s.filled]
        values = [v for v in values if v is not None]
        mixed = 1 if (any(v < 0 for v in values) and any(v > 0 for v in values)) else 0
        return (mixed, s.decimal_ratio, -s.index)

    numeric.sort(key=score, reverse=True)
    mapping.amount = numeric[0].index
    mapping.notes.append(
        "Flere talkolonner - valgte \"%s\" som beløb. Ret det i dialogen hvis det er forkert."
        % numeric[0].header)


def _pick_text(stats: Sequence[ColumnStats], used: Sequence[int]) -> List[int]:
    candidates = [s for s in stats
                  if s.index not in used and s.numeric_ratio < 0.5 and
                  s.date_ratio < 0.5 and s.fill_ratio > 0.2 and
                  not _has_word(s.header, IGNORE_WORDS)]
    if not candidates:
        return []

    def score(s: ColumnStats) -> tuple:
        named = 1 if _has_word(s.header, TEXT_WORDS) else 0
        return (named, s.distinct_ratio, min(s.avg_length, 40), -s.index)

    candidates.sort(key=score, reverse=True)
    best = candidates[0]
    chosen = [best.index]

    # Several named text columns (fx "Navn" + "Titel") describe one and the
    # same transaction - keep both, so the keyword rules have more to work on.
    if _has_word(best.header, TEXT_WORDS):
        for other in candidates[1:]:
            if len(chosen) >= 2:
                break
            if _has_word(other.header, TEXT_WORDS) and other.fill_ratio > 0.5 \
                    and other.avg_length >= 4:
                chosen.append(other.index)
    # A "type" column (few distinct values) alone is not descriptive enough.
    elif best.distinct_ratio < 0.25 and len(candidates) > 1:
        chosen.append(candidates[1].index)
    return sorted(chosen)


def detect_mapping(table, limit: int = 300) -> ColumnMapping:
    """Detect the roles of the columns of ``table``."""
    mapping = ColumnMapping()
    stats = analyse_columns(table, limit)
    if not stats:
        mapping.notes.append("Tabellen er tom.")
        return mapping

    mapping.date = _pick_date(stats)
    if mapping.date is not None:
        mapping.dayfirst = detect_dayfirst(stats[mapping.date].filled)

    numeric_values: List[str] = []
    for s in stats:
        if s.is_numeric() and s.index != mapping.date:
            numeric_values.extend(s.filled)
    mapping.decimal = detect_decimal_separator(numeric_values) or ","

    _pick_amount(stats, mapping.date, mapping)

    used = [i for i in (mapping.date, mapping.amount, mapping.amount_in,
                        mapping.amount_out, mapping.balance) if i is not None]
    for s in stats:
        if s.index not in used and _has_word(s.header, CURRENCY_WORDS) and \
                s.avg_length <= 6:
            mapping.currency = s.index
            used.append(s.index)
            break

    mapping.text = _pick_text(stats, used)
    if not mapping.text:
        mapping.notes.append("Fandt ingen tekstkolonne - alle poster får samme beskrivelse.")
    if mapping.date is None:
        mapping.notes.append("Fandt ingen datokolonne - vælg den manuelt.")
    return mapping
