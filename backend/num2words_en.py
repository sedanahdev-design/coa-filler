"""Minimal, dependency-free English number-to-words for USD amounts, matching
the "SAY <AMOUNT> USD ONLY." style seen on the commercial invoices (e.g.
30000 -> "THIRTY THOUSAND"). No network / no external package needed."""
from __future__ import annotations

_ONES = ["", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE",
         "TEN", "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN",
         "SEVENTEEN", "EIGHTEEN", "NINETEEN"]
_TENS = ["", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"]
_SCALES = [(1_000_000_000, "BILLION"), (1_000_000, "MILLION"), (1_000, "THOUSAND"), (100, "HUNDRED")]


def _three_digit_words(n: int) -> str:
    if n == 0:
        return ""
    words = []
    if n >= 100:
        words.append(_ONES[n // 100])
        words.append("HUNDRED")
        n %= 100
    if n >= 20:
        words.append(_TENS[n // 10])
        if n % 10:
            words.append(_ONES[n % 10])
    elif n > 0:
        words.append(_ONES[n])
    return " ".join(words)


def int_to_words(n: int) -> str:
    if n == 0:
        return "ZERO"
    if n < 0:
        return "MINUS " + int_to_words(-n)
    parts = []
    for scale_value, scale_name in [(1_000_000_000, "BILLION"), (1_000_000, "MILLION"), (1_000, "THOUSAND")]:
        if n >= scale_value:
            chunk = n // scale_value
            parts.append(f"{_three_digit_words(chunk)} {scale_name}")
            n %= scale_value
    if n > 0:
        parts.append(_three_digit_words(n))
    return " ".join(p for p in parts if p).strip()


def amount_to_words(amount: float, currency: str = "USD") -> str:
    """e.g. amount_to_words(30000) -> 'THIRTY THOUSAND USD ONLY.'"""
    whole = int(round(amount))
    return f"{int_to_words(whole)} {currency} ONLY."
