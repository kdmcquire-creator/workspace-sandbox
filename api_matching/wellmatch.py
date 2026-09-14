"""Shared helpers for matching Nexus header-file wells to Enverus well records."""

import re
from difflib import SequenceMatcher

# ---------------------------------------------------------------- normalization

_SPLIT_RE = re.compile(r"[^A-Z0-9]+")
_RUN_RE = re.compile(r"[0-9]+|[A-Z]+")


def norm_name(value):
    """Upper-case, strip punctuation, collapse whitespace."""
    if value is None:
        return ""
    return _SPLIT_RE.sub(" ", str(value).upper()).strip()


def norm_county(value):
    """'GRADY (OK)' and 'GRADY' both become 'GRADY'."""
    if value is None:
        return ""
    return re.sub(r"\s*\((?:OK|OKLAHOMA)\)\s*$", "", str(value).upper()).strip()


def canon_token(token):
    """Sort the letters inside each letter-run of a mixed alphanumeric token.

    Operators spell the horizontal designator inconsistently -- 7XHW, 7HXW and
    7WXH are the same well -- so the letter run is order-insensitive while the
    digits that identify the well stay put. Pure-word tokens are left alone:
    the lease name is not a permutation puzzle.
    """
    if not any(c.isdigit() for c in token):
        return token
    return "".join(
        run if run[0].isdigit() else "".join(sorted(run))
        for run in _RUN_RE.findall(token)
    )


def canon_name(value):
    """Normalized name with letter-runs sorted, for suffix-permutation matching."""
    return " ".join(canon_token(t) for t in norm_name(value).split())


def name_similarity(a, b):
    """Blend of whole-string and token-set similarity, 0..1."""
    if not a or not b:
        return 0.0
    seq = SequenceMatcher(None, a, b).ratio()
    ta, tb = set(a.split()), set(b.split())
    jaccard = len(ta & tb) / len(ta | tb) if (ta | tb) else 0.0
    return max(seq, (seq + jaccard) / 2)


# ---------------------------------------------------------------- field helpers


def to_float(value):
    try:
        f = float(value)
    except (TypeError, ValueError):
        return None
    return None if f != f else f


def iso_date(value):
    """First 10 chars of an ISO-ish date string, or None."""
    if value is None:
        return None
    s = str(value).strip()
    if len(s) < 10 or s.lower() in ("nan", "nat", "none", ""):
        return None
    return s[:10]


def digits(value):
    if value is None:
        return ""
    return re.sub(r"\D", "", str(value))
