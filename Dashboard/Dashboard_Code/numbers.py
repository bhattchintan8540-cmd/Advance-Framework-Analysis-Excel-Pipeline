"""Parsers for the dirty number and date formats in the raw extract."""

from __future__ import annotations

import re
from datetime import datetime

ONES = [
    "zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine",
    "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen",
    "seventeen", "eighteen", "nineteen",
]
TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]
WORD_VALUES = {word: index for index, word in enumerate(ONES)}
WORD_VALUES.update({name: index * 10 for index, name in enumerate(TENS) if name})
SCALES = {"hundred": 100, "thousand": 1000, "million": 1000000}

DATE_FORMATS = (
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%m/%d/%Y",
    "%m/%d/%y",
    "%d-%b-%Y",
    "%d-%b-%y",
    "%d.%m.%Y",
    "%d-%m-%Y",
    "%b %d, %Y",
    "%d %b %Y",
)


def int_to_words(value: int) -> str:
    """Spell an integer from 0 through 999,999. Used to build dirty source values."""
    if value < 0 or value >= 1_000_000:
        raise ValueError(f"int_to_words supports 0..999999, got {value}")

    def under_100(number: int) -> str:
        if number < 20:
            return ONES[number]
        tens, rest = divmod(number, 10)
        return TENS[tens] if rest == 0 else f"{TENS[tens]} {ONES[rest]}"

    def under_1000(number: int) -> str:
        if number < 100:
            return under_100(number)
        hundreds, rest = divmod(number, 100)
        if rest == 0:
            return f"{ONES[hundreds]} hundred"
        return f"{ONES[hundreds]} hundred {under_100(rest)}"

    if value < 1000:
        return under_1000(value)
    thousands, rest = divmod(value, 1000)
    if rest == 0:
        return f"{under_1000(thousands)} thousand"
    return f"{under_1000(thousands)} thousand {under_1000(rest)}"


def camel_words(value: int) -> str:
    return "".join(part.title() for part in int_to_words(value).split())


def words_to_number(text: str) -> float | None:
    raw = str(text).strip()
    if not re.search(r"[A-Za-z]", raw):
        return None
    spaced = re.sub(r"([A-Z])", r" \1", raw)
    spaced = re.sub(r"[^A-Za-z\s-]", " ", spaced)
    tokens = [token for token in re.split(r"[\s-]+", spaced.lower()) if token and token != "and"]
    if not tokens:
        return None
    if any(token not in WORD_VALUES and token not in SCALES for token in tokens):
        return None
    total = 0
    current = 0
    for token in tokens:
        if token in WORD_VALUES:
            current += WORD_VALUES[token]
        elif token == "hundred":
            current = (current or 1) * 100
        else:
            current = (current or 1) * SCALES[token]
            total += current
            current = 0
    return float(total + current)


def parse_number(value) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if text == "" or text.lower() in {"n/a", "na", "null", "none", "-", "nan"}:
        return None
    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()
    compact = text.replace("$", "").replace(",", "").replace(" ", "")
    if compact.endswith("%"):
        try:
            number = float(compact[:-1]) / 100.0
        except ValueError:
            number = None
        else:
            return -number if negative else number
    try:
        number = float(compact)
    except ValueError:
        number = words_to_number(str(value))
        if number is None:
            return None
    return -number if negative else number


def parse_date(value):
    if value is None:
        return None
    text = str(value).strip()
    if text == "" or text.lower() in {"n/a", "na", "null", "none"}:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None
