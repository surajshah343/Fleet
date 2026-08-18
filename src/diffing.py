"""Working out what was added and what disappeared.

Definitions used throughout:

* **New additions** -- rows whose key is present in the *newer* file but not in
  the older one.  All columns of the newer file are kept.
* **Deletions** -- rows whose key is present in the *older* file but not in the
  newer one.  All columns of the older file are kept.

The comparison is done on a normalised copy of the key columns, never on the
displayed values, so ``"978554    "`` matches ``"978554"``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence

import pandas as pd

from .cleaning import normalise_key

KEY_SEPARATOR = "\u241f"  # unit separator: cannot appear in real data


@dataclass
class DiffOptions:
    """How to compare the two sheets."""

    key_columns: Sequence[str]
    old_key_columns: Optional[Sequence[str]] = None
    uppercase: bool = True
    strip_leading_zeros: bool = False
    drop_missing_keys: bool = True


@dataclass
class DiffResult:
    """Additions, deletions and the counts behind them."""

    additions: pd.DataFrame
    deletions: pd.DataFrame
    n_new_rows: int = 0
    n_old_rows: int = 0
    n_common: int = 0
    n_duplicate_keys_new: int = 0
    n_duplicate_keys_old: int = 0
    n_missing_keys_new: int = 0
    n_missing_keys_old: int = 0
    duplicate_examples: List[str] = field(default_factory=list)

    def as_records(self) -> List[Dict[str, object]]:
        return [
            {"Measure": "Rows in newer file", "Value": self.n_new_rows},
            {"Measure": "Rows in older file", "Value": self.n_old_rows},
            {"Measure": "Assets in both files", "Value": self.n_common},
            {"Measure": "New additions", "Value": len(self.additions)},
            {"Measure": "Deletions", "Value": len(self.deletions)},
            {"Measure": "Duplicate keys (newer)", "Value": self.n_duplicate_keys_new},
            {"Measure": "Duplicate keys (older)", "Value": self.n_duplicate_keys_old},
            {"Measure": "Blank keys (newer)", "Value": self.n_missing_keys_new},
            {"Measure": "Blank keys (older)", "Value": self.n_missing_keys_old},
        ]


def _composite_key(
    frame: pd.DataFrame, columns: Sequence[str], options: DiffOptions
) -> pd.Series:
    """Join one or more normalised columns into a single comparison key."""
    missing = [c for c in columns if c not in frame.columns]
    if missing:
        raise KeyError(f"Key column(s) not found in sheet: {', '.join(missing)}")

    parts = [
        normalise_key(
            frame[column],
            uppercase=options.uppercase,
            strip_leading_zeros=options.strip_leading_zeros,
        )
        for column in columns
    ]
    key = parts[0].astype("string")
    for part in parts[1:]:
        key = key.str.cat(part.astype("string"), sep=KEY_SEPARATOR, na_rep="")
    return key


def compare_sheets(
    new_frame: pd.DataFrame, old_frame: pd.DataFrame, options: DiffOptions
) -> DiffResult:
    """Compare two cleaned sheets and return additions plus deletions."""
    old_columns = list(options.old_key_columns or options.key_columns)

    new_key = _composite_key(new_frame, options.key_columns, options)
    old_key = _composite_key(old_frame, old_columns, options)

    result = DiffResult(
        additions=new_frame.iloc[0:0].copy(),
        deletions=old_frame.iloc[0:0].copy(),
        n_new_rows=len(new_frame),
        n_old_rows=len(old_frame),
        n_missing_keys_new=int(new_key.isna().sum()),
        n_missing_keys_old=int(old_key.isna().sum()),
        n_duplicate_keys_new=int(new_key.dropna().duplicated().sum()),
        n_duplicate_keys_old=int(old_key.dropna().duplicated().sum()),
    )

    duplicated = new_key.dropna()[new_key.dropna().duplicated()].unique()[:5]
    result.duplicate_examples = [str(d).replace(KEY_SEPARATOR, " | ") for d in duplicated]

    new_valid = new_key.notna() if options.drop_missing_keys else pd.Series(
        True, index=new_frame.index
    )
    old_valid = old_key.notna() if options.drop_missing_keys else pd.Series(
        True, index=old_frame.index
    )

    old_set = set(old_key[old_valid].dropna())
    new_set = set(new_key[new_valid].dropna())

    added_mask = new_valid & ~new_key.isin(old_set)
    deleted_mask = old_valid & ~old_key.isin(new_set)

    additions = new_frame.loc[added_mask].copy()
    deletions = old_frame.loc[deleted_mask].copy()

    result.additions = additions.reset_index(drop=True)
    result.deletions = deletions.reset_index(drop=True)
    result.n_common = len(new_set & old_set)
    return result


def column_drift(
    new_columns: Sequence[str], old_columns: Sequence[str]
) -> pd.DataFrame:
    """Report columns that appear, disappear or are merely renamed."""
    from .naming import canonical

    new_map = {canonical(c): c for c in new_columns}
    old_map = {canonical(c): c for c in old_columns}

    rows: List[Dict[str, object]] = []
    for key, label in new_map.items():
        if key in old_map:
            if old_map[key] != label:
                rows.append(
                    {
                        "Column (newer file)": label,
                        "Column (older file)": old_map[key],
                        "Status": "Renamed / respaced",
                    }
                )
        else:
            rows.append(
                {
                    "Column (newer file)": label,
                    "Column (older file)": "",
                    "Status": "Only in newer file",
                }
            )
    for key, label in old_map.items():
        if key not in new_map:
            rows.append(
                {
                    "Column (newer file)": "",
                    "Column (older file)": label,
                    "Status": "Only in older file",
                }
            )
    return pd.DataFrame(rows, columns=["Column (newer file)", "Column (older file)", "Status"])
