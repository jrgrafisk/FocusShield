"""Find the recognisable shop or company inside a transaction text.

Bank exports rarely write just "Netto".  They write

    "Dankort-nota 4711 NETTO 8021 KØBENHAVN 21.05 kl. 17.42"
    "FØTEX 1234 – Dankort-nota 998877"
    "Visa kortkøb WOLT DANMARK 245,00 DKK"

Amounts, dates, times, card and receipt numbers make every single line
unique, so grouping on the raw text gives one group per transaction.  This
module strips that noise, so the lines above group as "NETTO KØBENHAVN",
"FØTEX" and "WOLT DANMARK" - and so a saved rule matches next month's
purchase in the same shop.
"""

from __future__ import annotations

import re
from typing import List

from .parsing import looks_like_amount
from .textutils import fold, squeeze

__all__ = ["merchant_name", "merchant_key", "rule_keyword"]

# Words that say something about the payment, not about who was paid.
NOISE_WORDS = frozenset((
    "dankort", "nota", "dankortnota", "kortkob", "kortkoeb", "kort", "kortnr",
    "visa", "mastercard", "maestro", "eurocard", "kreditkort", "betalingskort",
    "debitkort", "chipkort", "haevekort", "kontokort", "korttype",
    "kortbetaling", "korttransaktion", "kortnota", "kortholder",
    "kortoplysninger", "kortkobsnota",
    "bs", "pbs", "betaling", "betalingsservice", "indbetalingskort",
    "ref", "refnr", "reference", "referencenr", "id", "idnr", "nr", "no",
    "kl", "den", "d", "dato", "tid", "kvittering", "bilag", "faktura", "fakt",
    "kob", "koeb", "purchase", "payment", "pos", "atm", "automat", "netbank",
    "mobilbank", "onlinebank", "udland", "udl", "valutakurs", "kurs",
    "dk", "dkk", "sek", "nok", "eur", "usd", "gbp", "kr", "kroner",
    "fra", "til", "med", "m", "v", "og", "af", "pa", "konto", "kontonr",
    "konto-nr",
    "transaktion", "postering", "posteringstekst", "tekst", "note",
    "aut", "auto", "automatisk", "straks", "straksoverforsel",
))

# Tokens that are a date ("21.05", "2025-05-21") or a time ("17.42", "17:42").
_DATE_RE = re.compile(r"^\d{1,4}[./-]\d{1,2}([./-]\d{2,4})?\.?$")
_TIME_RE = re.compile(r"^\d{1,2}[:.]\d{2}$")
_LETTER_RE = re.compile(r"[^\W\d_]", re.UNICODE)
_TRIM_RE = re.compile(r"^[^\w&]+|[^\w&%]+$", re.UNICODE)

# How much of a merchant name a saved rule should use.
RULE_WORDS = 3
RULE_CHARS = 32
MIN_NAME_LENGTH = 3


def _is_noise(token: str) -> bool:
    """True when a token says something about the payment, not the payee."""
    if not _LETTER_RE.search(token):
        return True                         # pure numbers, "//", "-", "***"
    if _DATE_RE.match(token) or _TIME_RE.match(token):
        return True
    if looks_like_amount(token):
        return True                         # "245,00", "1.234,56", "99kr"
    folded = fold(token)
    if folded in NOISE_WORDS:
        return True
    # "Dankort-nota", "kortkøb/visa": noise when every part is noise.
    parts = [part for part in re.split(r"[-/.]", folded) if part]
    if len(parts) > 1 and all(part in NOISE_WORDS for part in parts):
        return True
    digits = sum(1 for char in token if char.isdigit())
    if digits and digits * 2 >= len(token):
        return True                         # "8021", "4711a", "12-34"
    return False


def _is_connector(token: str) -> bool:
    """"&", "-", "/" and friends: part of a name, not a word of their own."""
    return len(token) <= 2 and not _LETTER_RE.search(token) and \
        not any(char.isdigit() for char in token)


def _analyse(text: object):
    """``(tokens, keep_flags, connector_flags)`` for the words of ``text``."""
    tokens = [t for t in squeeze(text).split(" ") if t]
    keep = []
    connectors = []
    for token in tokens:
        trimmed = _TRIM_RE.sub("", token) or token
        connector = _is_connector(token)
        connectors.append(connector)
        keep.append(not connector and not _is_noise(trimmed))
    return tokens, keep, connectors


def _clean_tokens(text: object) -> List[str]:
    """The tokens of ``text`` that actually name someone."""
    tokens, keep, _connectors = _analyse(text)
    return [_TRIM_RE.sub("", token) or token
            for token, wanted in zip(tokens, keep) if wanted]


def merchant_name(text: object) -> str:
    """The shop/company part of a transaction text (original spelling)."""
    tokens = _clean_tokens(text)
    name = " ".join(tokens).strip()
    if len(name) < MIN_NAME_LENGTH:
        return squeeze(text)
    return name


def merchant_key(text: object) -> str:
    """A grouping key: two texts from the same shop give the same key."""
    return fold(merchant_name(text)) or fold(text)


def rule_keyword(text: object) -> str:
    """A keyword to save as a rule - short enough to match next month too.

    Rules match as plain substrings, so the keyword has to be an unbroken
    piece of the original text: "JOE & THE JUICE" must keep its "&".
    """
    tokens, keep, connectors = _analyse(text)
    if not any(keep):
        return squeeze(text)[:RULE_CHARS]
    first = keep.index(True)
    last = first
    words = 0
    for index in range(first, len(tokens)):
        if connectors[index]:
            continue                        # "&" does not end the name
        if not keep[index]:
            break                           # a shop number ends it: "NETTO 8021"
        words += 1
        if words > RULE_WORDS:
            break
        last = index
    keyword = _TRIM_RE.sub("", " ".join(tokens[first:last + 1]))
    if len(keyword) > RULE_CHARS:
        keyword = keyword[:RULE_CHARS].rstrip()
    return keyword or squeeze(text)[:RULE_CHARS]
