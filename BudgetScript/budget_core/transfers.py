"""Recognise transactions that move money between the user's own accounts.

Bank exports have no reliable flag for "this is not real income" - a
transfer from a savings account into the checking account looks exactly
like a salary payment.  The only thing that gives it away is the account
number written in the text ("Overført fra 1234567890",
"Til konto 5301-1234567890").  If the user tells us which account numbers
are their own (on the Konti sheet), we can catch those and keep them out
of the budget entirely.
"""

from __future__ import annotations

import re
from typing import List, Sequence

from .textutils import squeeze

__all__ = ["MIN_ACCOUNT_DIGITS", "normalize_account_number",
          "parse_account_numbers", "text_mentions_account"]

# Danish account numbers run to 10 digits (plus a 4 digit reg. number).
# Anything shorter is too likely to collide with a date or a receipt number,
# so we simply never match on it.
MIN_ACCOUNT_DIGITS = 6

_SPLIT_RE = re.compile(r"[,;/]+")


def normalize_account_number(raw: object) -> str:
    """Strip everything but digits: ``"1234 5678901"`` -> ``"12345678901"``."""
    return "".join(ch for ch in str(raw or "") if ch.isdigit())


def parse_account_numbers(raw: object) -> List[str]:
    """The account numbers in one Konti cell (comma/semicolon separated)."""
    if not raw:
        return []
    numbers = []
    for part in _SPLIT_RE.split(str(raw)):
        digits = normalize_account_number(part)
        if len(digits) >= MIN_ACCOUNT_DIGITS:
            numbers.append(digits)
    return numbers


def text_mentions_account(text: object, account_numbers: Sequence[str]) -> bool:
    """True when ``text`` contains one of ``account_numbers``.

    Matches a single token whose digits contain the account number (so
    "5301-1234567890" and plain "1234567890" both match), and also two
    adjacent pure-digit tokens concatenated (so "5301 1234567890",
    reg. number and account number split by a space, matches too).
    """
    numbers = [n for n in account_numbers if n]
    if not numbers:
        return False
    tokens = squeeze(text).split(" ")
    token_digits = [normalize_account_number(t) for t in tokens]
    for digits in token_digits:
        if len(digits) < MIN_ACCOUNT_DIGITS:
            continue
        if any(number in digits for number in numbers):
            return True
    for i in range(len(tokens) - 1):
        if tokens[i].isdigit() and tokens[i + 1].isdigit():
            combined = tokens[i] + tokens[i + 1]
            if any(number in combined for number in numbers):
                return True
    return False
