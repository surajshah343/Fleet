"""Data cleansing.

Every step is optional and every step reports what it touched, so the app can
show a before/after audit instead of silently rewriting the user's data.

The problems seen in the sample workbooks:

* trailing padding on almost every text column (``"MAX ATLAS        "``)
* tabs and non-breaking spaces inside cells
* placeholder nulls: ``"NULL"``, ``"--"``, ``"N/A"``, ``" "``, ``"#N/A"``
* money and counts stored as text (``FAIRMARKETVALUE`` reads as a string)
* dates stored three ways in the same column family: real datetimes, ISO
  strings, and ``1900-01-01`` used as a stand-in for "no date"
* fully blank spacer rows and columns
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import numpy as np
import pandas as pd

from .naming import strip_invisibles

NULL_TOKENS = {
    "", "-", "--", "---", "n/a", "n\\a", "na", "null", "none", "nil",
    "#n/a", "#value!", "#ref!", "#div/0!", "#name?", "#null!", "#num!",
    "nan", "nat", "not applicable", "unknown", "tbd", ".",
}
"""Lower-cased strings treated as missing values."""

PLACEHOLDER_DATES = {"1900-01-01", "1899-12-30", "1970-01-01"}
"""Epoch-ish dates Excel exports use to mean "no date"."""

MIN_SAMPLE_FOR_TYPING = 3
"""Never re-type a column from one or two stray values."""

_NUMERIC_PATTERN = re.compile(
    r"^\(?\s*[-+]?\s*[$€£]?\s*\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*%?\s*\)?$|"
    r"^\(?\s*[-+]?\s*[$€£]?\s*\d*\.?\d+(?:[eE][-+]?\d+)?\s*%?\s*\)?$"
)

_DATE_HINT = re.compile(r"date|expir|dob|\bday\b|month|year$", re.IGNORECASE)


@dataclass
class CleaningOptions:
    """Toggles for the cleaning pipeline."""

    trim_whitespace: bool = True
    collapse_inner_spaces: bool = True
    normalise_nulls: bool = True
    drop_blank_rows: bool = True
    drop_blank_columns: bool = True
    convert_numeric: bool = True
    convert_dates: bool = True
    blank_placeholder_dates: bool = True
    drop_duplicate_rows: bool = False
    uppercase_keys: bool = False
    numeric_threshold: float = 0.95
    date_threshold: float = 0.90


@dataclass
class CleaningReport:
    """What the pipeline actually changed."""

    rows_before: int = 0
    rows_after: int = 0
    columns_before: int = 0
    columns_after: int = 0
    cells_trimmed: int = 0
    cells_nulled: int = 0
    blank_rows_dropped: int = 0
    duplicate_rows_dropped: int = 0
    columns_dropped: List[str] = field(default_factory=list)
    converted_to_numeric: List[str] = field(default_factory=list)
    converted_to_datetime: List[str] = field(default_factory=list)
    placeholder_dates_blanked: Dict[str, int] = field(default_factory=dict)

    def as_records(self) -> List[Dict[str, object]]:
        """Flatten into rows for display in a table."""
        return [
            {"Check": "Rows in", "Result": self.rows_before},
            {"Check": "Rows out", "Result": self.rows_after},
            {"Check": "Columns in", "Result": self.columns_before},
            {"Check": "Columns out", "Result": self.columns_after},
            {"Check": "Cells trimmed of whitespace", "Result": self.cells_trimmed},
            {"Check": "Placeholder values set to blank", "Result": self.cells_nulled},
            {"Check": "Blank rows removed", "Result": self.blank_rows_dropped},
            {"Check": "Duplicate rows removed", "Result": self.duplicate_rows_dropped},
            {"Check": "Empty columns removed", "Result": len(self.columns_dropped)},
            {"Check": "Text columns turned numeric", "Result": len(self.converted_to_numeric)},
            {"Check": "Text columns turned into dates", "Result": len(self.converted_to_datetime)},
            {
                "Check": "Placeholder dates blanked",
                "Result": int(sum(self.placeholder_dates_blanked.values())),
            },
        ]


def _is_texty(series: pd.Series) -> bool:
    return series.dtype == object or str(series.dtype) in {"str", "string"}


def _clean_text_series(series: pd.Series, options: CleaningOptions):
    """Trim/collapse a text column. Returns (series, n_trimmed, n_nulled)."""
    original = series.astype("string")
    working = original

    if options.trim_whitespace or options.collapse_inner_spaces:
        working = working.map(
            lambda v: strip_invisibles(v) if isinstance(v, str) else v,
            na_action="ignore",
        )
    if options.collapse_inner_spaces:
        working = working.str.replace(r"\s+", " ", regex=True)
    if options.trim_whitespace:
        working = working.str.strip()

    trimmed = int((original.fillna("") != working.fillna("")).sum())

    nulled = 0
    if options.normalise_nulls:
        lowered = working.str.lower().str.strip()
        mask = lowered.isin(NULL_TOKENS) & working.notna()
        nulled = int(mask.sum())
        working = working.mask(mask, pd.NA)

    return working, trimmed, nulled


def _try_numeric(series: pd.Series, threshold: float) -> Optional[pd.Series]:
    """Convert a text column to numbers when nearly all values qualify."""
    text = series.dropna().astype(str).str.strip()
    if len(text) < MIN_SAMPLE_FOR_TYPING:
        return None
    if not (text.str.match(_NUMERIC_PATTERN).mean() >= threshold):
        return None

    cleaned = series.astype("string").str.strip()
    negative = cleaned.str.match(r"^\(.*\)$", na=False)
    percent = cleaned.str.endswith("%", na=False)
    stripped = cleaned.str.replace(r"[,$€£()%\s]", "", regex=True)

    numbers = pd.to_numeric(stripped, errors="coerce")
    numbers = numbers.where(~negative, -numbers.abs())
    numbers = numbers.where(~percent, numbers / 100.0)
    return numbers


def _try_datetime(name: str, series: pd.Series, threshold: float) -> Optional[pd.Series]:
    """Convert a text column to datetimes when the name and values agree."""
    text = series.dropna().astype(str).str.strip()
    if len(text) < MIN_SAMPLE_FOR_TYPING:
        return None

    sample = text.head(400)
    looks_dateish = sample.str.match(
        r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}|^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}", na=False
    ).mean()
    if looks_dateish < threshold and not _DATE_HINT.search(name):
        return None

    converted = pd.to_datetime(series, errors="coerce", format="mixed")
    if converted.notna().sum() / max(series.notna().sum(), 1) < threshold:
        return None
    return converted


def clean_dataframe(
    frame: pd.DataFrame, options: Optional[CleaningOptions] = None
) -> tuple[pd.DataFrame, CleaningReport]:
    """Run the cleaning pipeline. Returns the cleaned copy and a report."""
    options = options or CleaningOptions()
    report = CleaningReport(
        rows_before=len(frame), columns_before=frame.shape[1]
    )
    out = frame.copy()

    # 1. Text hygiene ------------------------------------------------------
    for column in out.columns:
        if _is_texty(out[column]):
            cleaned, trimmed, nulled = _clean_text_series(out[column], options)
            out[column] = cleaned
            report.cells_trimmed += trimmed
            report.cells_nulled += nulled

    # 2. Structural drops --------------------------------------------------
    if options.drop_blank_rows:
        blank = out.isna().all(axis=1)
        report.blank_rows_dropped = int(blank.sum())
        out = out.loc[~blank]

    if options.drop_blank_columns:
        empty = [c for c in out.columns if out[c].isna().all()]
        report.columns_dropped = empty
        out = out.drop(columns=empty)

    if options.drop_duplicate_rows:
        before = len(out)
        out = out.drop_duplicates()
        report.duplicate_rows_dropped = before - len(out)

    # 3. Type recovery -----------------------------------------------------
    if options.convert_dates:
        for column in list(out.columns):
            if _is_texty(out[column]):
                converted = _try_datetime(column, out[column], options.date_threshold)
                if converted is not None:
                    out[column] = converted
                    report.converted_to_datetime.append(column)

    if options.convert_numeric:
        for column in list(out.columns):
            if _is_texty(out[column]) and column not in report.converted_to_datetime:
                converted = _try_numeric(out[column], options.numeric_threshold)
                if converted is not None:
                    out[column] = converted
                    report.converted_to_numeric.append(column)

    # 4. Placeholder dates -------------------------------------------------
    if options.blank_placeholder_dates:
        placeholders = pd.to_datetime(sorted(PLACEHOLDER_DATES))
        for column in out.columns:
            if pd.api.types.is_datetime64_any_dtype(out[column]):
                mask = out[column].isin(placeholders)
                count = int(mask.sum())
                if count:
                    out[column] = out[column].mask(mask)
                    report.placeholder_dates_blanked[column] = count

    out = out.reset_index(drop=True)
    report.rows_after = len(out)
    report.columns_after = out.shape[1]
    return out, report


def normalise_key(series: pd.Series, uppercase: bool = True,
                  strip_leading_zeros: bool = False) -> pd.Series:
    """Build a comparable key from an identifier column.

    Removes *all* whitespace (serials arrive as ``"978554    "`` in one file and
    ``"978554"`` in the other) and optionally case and leading zeros.
    """
    key = series.astype("string").map(
        lambda v: strip_invisibles(v) if isinstance(v, str) else v,
        na_action="ignore",
    )
    key = key.str.replace(r"\s+", "", regex=True)
    if uppercase:
        key = key.str.upper()
    if strip_leading_zeros:
        key = key.str.replace(r"^0+(?=.)", "", regex=True)
    return key.mask(key.str.lower().isin(NULL_TOKENS) | (key == ""), pd.NA)


def suggest_key_columns(columns: List[str]) -> List[str]:
    """Rank likely identifier columns so the UI can pre-select a sensible key."""
    from .naming import canonical

    preference = [
        "SERIALNUMBER", "VIN", "UNITNUMBER", "UNIT", "ASSETID", "ASSETNUMBER",
        "EQUIPMENTID", "HILCOREF",
    ]
    ranked: List[str] = []
    for wanted in preference:
        for column in columns:
            if canonical(column) == wanted and column not in ranked:
                ranked.append(column)
    return ranked
