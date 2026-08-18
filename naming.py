"""Name normalisation for columns and sheets.

The two workbooks do not agree on spelling.  The US sheets use
``UnitNumber`` / ``NBV_Current`` while the Canadian sheets use
``Unit Number`` / ``NBV Current``.  Sheet names differ too
(``USA``/``Canada`` in one file, ``US``/``CAN`` in the other).

Everything in this module maps a messy label onto a canonical key so the
two files can be lined up.  Output always keeps the *original* labels --
canonical keys are only ever used for matching.
"""

from __future__ import annotations

import re
import unicodedata
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

# Characters Excel exports love to hide inside cells and headers.
_INVISIBLES = dict.fromkeys(
    map(ord, "\u00a0\u1680\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007"
             "\u2008\u2009\u200a\u200b\u202f\u205f\u3000\ufeff\u200c\u200d"),
    " ",
)


def strip_invisibles(text: str) -> str:
    """Replace non-breaking spaces, zero-width chars and tabs with spaces."""
    text = unicodedata.normalize("NFKC", text)
    text = text.translate(_INVISIBLES)
    return text.replace("\t", " ").replace("\r", " ").replace("\n", " ")


def tidy_label(label: object) -> str:
    """Human-facing tidy-up: collapse whitespace runs, trim the ends."""
    if label is None:
        return ""
    text = strip_invisibles(str(label))
    return re.sub(r"\s+", " ", text).strip()


def canonical(label: object) -> str:
    """Aggressive key used only for matching.

    ``"Unit Number"``, ``"UnitNumber"`` and ``"unit_number"`` all collapse to
    ``"UNITNUMBER"``.
    """
    text = tidy_label(label).upper()
    return re.sub(r"[^0-9A-Z]", "", text)


def dedupe_labels(labels: Iterable[object]) -> List[str]:
    """Tidy labels and make them unique, Excel-style (``Name``, ``Name_2``)."""
    seen: Dict[str, int] = {}
    out: List[str] = []
    for index, raw in enumerate(labels):
        label = tidy_label(raw)
        if not label or label.lower().startswith("unnamed:"):
            label = f"Column_{index + 1}"
        seen[label] = seen.get(label, 0) + 1
        out.append(label if seen[label] == 1 else f"{label}_{seen[label]}")
    return out


# --------------------------------------------------------------------------
# Sheet pairing
# --------------------------------------------------------------------------

# Known aliases between the two workbook generations.
_SHEET_ALIASES: Dict[str, str] = {
    "US": "US",
    "USA": "US",
    "UNITEDSTATES": "US",
    "CAN": "CANADA",
    "CANADA": "CANADA",
    "CA": "CANADA",
    "SUMMARY": "SUMMARY",
    "HILCOSUMMARY": "SUMMARY",
}


def sheet_family(sheet_name: str) -> str:
    """Map a sheet name onto a region family (``US``, ``CANADA``, ...)."""
    key = canonical(sheet_name)
    if key in _SHEET_ALIASES:
        return _SHEET_ALIASES[key]
    for alias, family in _SHEET_ALIASES.items():
        if key.startswith(alias) or alias.startswith(key):
            return family
    return key


def pair_sheets(
    new_sheets: Sequence[str], old_sheets: Sequence[str]
) -> List[Tuple[str, Optional[str]]]:
    """Suggest ``(new_sheet, old_sheet)`` pairs, best guess first.

    Unmatched new sheets come back paired with ``None`` so the UI can ask the
    user to resolve them by hand.
    """
    remaining = list(old_sheets)
    pairs: List[Tuple[str, Optional[str]]] = []
    for new_sheet in new_sheets:
        family = sheet_family(new_sheet)
        match = next((s for s in remaining if sheet_family(s) == family), None)
        if match is not None:
            remaining.remove(match)
        pairs.append((new_sheet, match))
    return pairs


def build_column_map(
    new_columns: Sequence[str], old_columns: Sequence[str]
) -> Dict[str, Optional[str]]:
    """Map each new-file column onto its old-file equivalent (or ``None``)."""
    old_by_key = {canonical(c): c for c in old_columns}
    return {col: old_by_key.get(canonical(col)) for col in new_columns}
