"""Tests for the comparison pipeline.

Run with:  python -m pytest tests/ -v
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.calc import (  # noqa: E402
    FormulaError,
    Operation,
    apply_operations,
    build_binary_formula,
    build_multi_formula,
    evaluate_formula,
)
from src.cleaning import CleaningOptions, clean_dataframe, normalise_key  # noqa: E402
from src.diffing import DiffOptions, column_drift, compare_sheets  # noqa: E402
from src.exporter import SheetPayload, append_sheets, safe_sheet_name, write_new_workbook  # noqa: E402
from src.loader import detect_header_row, read_sheet  # noqa: E402
from src.naming import canonical, build_column_map, pair_sheets  # noqa: E402


# --------------------------------------------------------------------------
# Fixtures
# --------------------------------------------------------------------------

@pytest.fixture
def messy_frame() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "SerialNumber": ["  ABC123 ", "def456", "GHI789\t", "  ", "JKL012"],
            "UnitNumber": ["001  ", "002", "003", "004", "005"],
            "OEM": ["MAX ATLAS      ", "NULL", "GREAT  DANE", "--", "WABASH"],
            "MonthlyLeaseRate": ["1,200.50", "$800", "(150)", "N/A", "0"],
            "InServiceDate": ["2023-07-26", "2024-01-15", "1900-01-01", "", "2025-03-01"],
            "Blank": [None, None, None, None, None],
        }
    )


@pytest.fixture
def old_frame() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Serial Number": ["ABC123", "DEF456", "ZZZ999"],
            "Unit Number": ["001", "002", "099"],
            "OEM": ["MAX ATLAS", "STRICK", "UTILITY"],
        }
    )


# --------------------------------------------------------------------------
# Naming
# --------------------------------------------------------------------------

def test_canonical_collapses_spelling_variants():
    assert canonical("Unit Number") == canonical("UnitNumber") == canonical("unit_number")
    assert canonical("NBV_Current") == canonical("NBV Current")


def test_pair_sheets_matches_regional_aliases():
    pairs = dict(pair_sheets(["USA", "Canada", "Hilco Summary"], ["US", "CAN", "Summary"]))
    assert pairs == {"USA": "US", "Canada": "CAN", "Hilco Summary": "Summary"}


def test_build_column_map_bridges_naming_styles():
    mapping = build_column_map(["UnitNumber", "Brand New"], ["Unit Number", "Other"])
    assert mapping["UnitNumber"] == "Unit Number"
    assert mapping["Brand New"] is None


# --------------------------------------------------------------------------
# Cleaning
# --------------------------------------------------------------------------

def test_cleaning_trims_and_nulls(messy_frame):
    cleaned, report = clean_dataframe(messy_frame)
    assert cleaned.loc[0, "SerialNumber"] == "ABC123"
    assert cleaned.loc[2, "OEM"] == "GREAT DANE"      # inner run collapsed
    assert pd.isna(cleaned.loc[1, "OEM"])             # "NULL" blanked
    assert pd.isna(cleaned.loc[3, "OEM"])             # "--" blanked
    assert report.cells_trimmed > 0
    assert report.cells_nulled >= 2


def test_cleaning_recovers_numeric_types(messy_frame):
    cleaned, report = clean_dataframe(messy_frame)
    assert "MonthlyLeaseRate" in report.converted_to_numeric
    assert cleaned.loc[0, "MonthlyLeaseRate"] == pytest.approx(1200.50)
    assert cleaned.loc[1, "MonthlyLeaseRate"] == pytest.approx(800.0)
    # Parentheses are accounting notation for a negative number.
    assert cleaned.loc[2, "MonthlyLeaseRate"] == pytest.approx(-150.0)


def test_cleaning_blanks_placeholder_dates(messy_frame):
    cleaned, report = clean_dataframe(messy_frame)
    assert "InServiceDate" in report.converted_to_datetime
    assert pd.isna(cleaned.loc[2, "InServiceDate"])   # 1900-01-01 is not a real date
    assert cleaned.loc[0, "InServiceDate"] == pd.Timestamp("2023-07-26")


def test_cleaning_drops_empty_columns(messy_frame):
    cleaned, report = clean_dataframe(messy_frame)
    assert "Blank" not in cleaned.columns
    assert "Blank" in report.columns_dropped


def test_cleaning_can_be_switched_off(messy_frame):
    options = CleaningOptions(
        trim_whitespace=False, collapse_inner_spaces=False, normalise_nulls=False,
        convert_numeric=False, convert_dates=False, drop_blank_columns=False,
        blank_placeholder_dates=False,
    )
    cleaned, _ = clean_dataframe(messy_frame, options)
    assert cleaned.loc[0, "SerialNumber"] == "  ABC123 "
    assert "Blank" in cleaned.columns


def test_normalise_key_ignores_padding_and_case():
    keys = normalise_key(pd.Series(["978554    ", "978554", " 978554"]))
    assert keys.nunique() == 1
    assert normalise_key(pd.Series(["abc"]))[0] == "ABC"


# --------------------------------------------------------------------------
# Diffing
# --------------------------------------------------------------------------

def test_compare_finds_additions_and_deletions(messy_frame, old_frame):
    new_clean, _ = clean_dataframe(messy_frame)
    old_clean, _ = clean_dataframe(old_frame)

    result = compare_sheets(
        new_clean,
        old_clean,
        DiffOptions(key_columns=["SerialNumber"], old_key_columns=["Serial Number"]),
    )
    added = set(result.additions["SerialNumber"])
    deleted = set(result.deletions["Serial Number"])

    assert added == {"GHI789", "JKL012"}   # ABC123/DEF456 exist in both
    assert deleted == {"ZZZ999"}
    assert result.n_common == 2


def test_compare_keeps_every_column(messy_frame, old_frame):
    new_clean, _ = clean_dataframe(messy_frame)
    old_clean, _ = clean_dataframe(old_frame)
    result = compare_sheets(
        new_clean, old_clean,
        DiffOptions(key_columns=["SerialNumber"], old_key_columns=["Serial Number"]),
    )
    assert list(result.additions.columns) == list(new_clean.columns)
    assert list(result.deletions.columns) == list(old_clean.columns)


def test_compare_is_case_and_space_insensitive():
    new = pd.DataFrame({"Serial": ["abc 123", "NEW1"]})
    old = pd.DataFrame({"Serial": ["ABC123"]})
    result = compare_sheets(new, old, DiffOptions(key_columns=["Serial"]))
    assert list(result.additions["Serial"]) == ["NEW1"]
    assert result.deletions.empty


def test_compare_supports_composite_keys():
    new = pd.DataFrame({"A": ["1", "1"], "B": ["x", "y"]})
    old = pd.DataFrame({"A": ["1"], "B": ["x"]})
    result = compare_sheets(new, old, DiffOptions(key_columns=["A", "B"]))
    assert len(result.additions) == 1
    assert result.additions.iloc[0]["B"] == "y"


def test_compare_reports_blank_and_duplicate_keys():
    new = pd.DataFrame({"S": ["A", "A", None, "B"]})
    old = pd.DataFrame({"S": ["A"]})
    result = compare_sheets(new, old, DiffOptions(key_columns=["S"]))
    assert result.n_duplicate_keys_new == 1
    assert result.n_missing_keys_new == 1
    assert list(result.additions["S"]) == ["B"]


def test_missing_key_column_raises():
    with pytest.raises(KeyError):
        compare_sheets(
            pd.DataFrame({"A": [1]}), pd.DataFrame({"A": [1]}),
            DiffOptions(key_columns=["Nope"]),
        )


def test_column_drift_flags_renames_and_new_columns():
    drift = column_drift(["UnitNumber", "Extra"], ["Unit Number", "Gone"])
    statuses = dict(zip(drift["Column (newer file)"], drift["Status"]))
    assert statuses["UnitNumber"] == "Renamed / respaced"
    assert statuses["Extra"] == "Only in newer file"
    assert "Only in older file" in set(drift["Status"])


# --------------------------------------------------------------------------
# Calculations
# --------------------------------------------------------------------------

@pytest.fixture
def calc_frame() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Rental Rate": [100.0, 200.0, 300.0],
            "Lease Rate": [50.0, 25.0, 0.0],
            "Model Year": [2019, 2021, 2024],
        }
    )


def test_binary_builder_adds_two_columns(calc_frame):
    formula = build_binary_formula("Add (+)", "[Rental Rate]", "[Lease Rate]")
    assert list(evaluate_formula(calc_frame, formula)) == [150.0, 225.0, 300.0]


def test_builder_supports_every_listed_operator(calc_frame):
    from src.calc import BUILDER_OPERATORS

    for operator in BUILDER_OPERATORS:
        formula = build_binary_formula(operator, "[Rental Rate]", "[Lease Rate]")
        result = evaluate_formula(calc_frame, formula)
        assert len(result) == len(calc_frame), operator


def test_multi_column_operator(calc_frame):
    formula = build_multi_formula("Sum of columns", ["Rental Rate", "Lease Rate"])
    assert list(evaluate_formula(calc_frame, formula)) == [150.0, 225.0, 300.0]


def test_operations_chain_onto_earlier_results(calc_frame):
    result = apply_operations(
        calc_frame,
        [
            Operation("Total", "[Rental Rate] + [Lease Rate]"),
            Operation("Annual", "[Total] * 12"),
        ],
    )
    assert result.applied == ["Total", "Annual"]
    assert list(result.frame["Annual"]) == [1800.0, 2700.0, 3600.0]


def test_rounding_option(calc_frame):
    result = apply_operations(
        calc_frame, [Operation("Third", "[Rental Rate] / 3", round_to=2)]
    )
    assert result.frame["Third"].iloc[0] == pytest.approx(33.33)


def test_division_by_zero_becomes_blank(calc_frame):
    result = apply_operations(calc_frame, [Operation("Ratio", "[Rental Rate] / [Lease Rate]")])
    assert pd.isna(result.frame["Ratio"].iloc[2])


def test_column_aggregate_broadcasts(calc_frame):
    values = evaluate_formula(calc_frame, "[Rental Rate] / COLSUM([Rental Rate]) * 100")
    assert values.sum() == pytest.approx(100.0)


def test_conditional_and_text_functions(calc_frame):
    values = evaluate_formula(calc_frame, 'IF([Model Year] >= 2021, "new", "old")')
    assert list(values) == ["old", "new", "new"]


def test_numeric_text_column_is_coerced():
    frame = pd.DataFrame({"Amount": ["1,000", "$2,500", "(300)"]})
    assert list(evaluate_formula(frame, "[Amount] * 2")) == [2000.0, 5000.0, -600.0]


def test_unknown_column_is_rejected(calc_frame):
    with pytest.raises(FormulaError):
        evaluate_formula(calc_frame, "[Nope] + 1")


@pytest.mark.parametrize(
    "formula",
    [
        '__import__("os").system("ls")',
        "open('/etc/passwd').read()",
        "[Rental Rate].__class__",
        "lambda: 1",
        "[x for x in range(3)]",
    ],
)
def test_sandbox_rejects_dangerous_formulas(calc_frame, formula):
    with pytest.raises((FormulaError, SyntaxError)):
        evaluate_formula(calc_frame, formula)


def test_errors_are_collected_not_raised(calc_frame):
    result = apply_operations(
        calc_frame,
        [Operation("Good", "[Rental Rate] * 2"), Operation("Bad", "[Missing] * 2")],
    )
    assert result.applied == ["Good"]
    assert result.errors[0][0] == "Bad"


# --------------------------------------------------------------------------
# Export
# --------------------------------------------------------------------------

def test_safe_sheet_name_strips_illegal_characters():
    assert safe_sheet_name("Sales/2026[Q1]") == "Sales 2026 Q1"
    assert len(safe_sheet_name("x" * 60)) == 31
    assert safe_sheet_name("Data", taken=["data"]) != "Data"


def test_new_workbook_round_trip(tmp_path, calc_frame):
    out = tmp_path / "out.xlsx"
    write_new_workbook(out, [SheetPayload("Results", calc_frame)])
    loaded = pd.read_excel(out, sheet_name="Results")
    pd.testing.assert_frame_equal(loaded, calc_frame, check_dtype=False)


def test_append_preserves_original_sheets(tmp_path, calc_frame):
    source = tmp_path / "source.xlsx"
    original = pd.DataFrame({"Keep": [1, 2, 3], "Text": ["a", "b", "c"]})
    write_new_workbook(source, [SheetPayload("Original", original)])

    destination = tmp_path / "appended.xlsx"
    append_sheets(source, destination, [SheetPayload("New Additions", calc_frame)])

    sheets = pd.read_excel(destination, sheet_name=None)
    assert set(sheets) == {"Original", "New Additions"}
    pd.testing.assert_frame_equal(sheets["Original"], original, check_dtype=False)
    pd.testing.assert_frame_equal(sheets["New Additions"], calc_frame, check_dtype=False)


def test_append_handles_dates_and_blanks(tmp_path):
    frame = pd.DataFrame(
        {
            "When": pd.to_datetime(["2026-06-19", None, "2025-01-02"]),
            "Value": [1.5, None, 3.0],
            "Label": ["a", None, "c"],
        }
    )
    source = tmp_path / "s.xlsx"
    write_new_workbook(source, [SheetPayload("S", pd.DataFrame({"x": [1]}))])
    out = tmp_path / "d.xlsx"
    append_sheets(source, out, [SheetPayload("Dates", frame)])

    loaded = pd.read_excel(out, sheet_name="Dates")
    assert loaded["When"].iloc[0] == pd.Timestamp("2026-06-19")
    assert pd.isna(loaded["When"].iloc[1])
    assert loaded["Value"].iloc[2] == 3.0
    assert pd.isna(loaded["Label"].iloc[1])


def test_append_replaces_a_sheet_of_the_same_name(tmp_path):
    source = tmp_path / "s.xlsx"
    write_new_workbook(
        source,
        [SheetPayload("Keep", pd.DataFrame({"a": [1]})),
         SheetPayload("Additions", pd.DataFrame({"old": [1, 2, 3]}))],
    )
    out = tmp_path / "d.xlsx"
    append_sheets(source, out, [SheetPayload("Additions", pd.DataFrame({"new": [9]}))])

    sheets = pd.read_excel(out, sheet_name=None)
    assert set(sheets) == {"Keep", "Additions"}
    assert list(sheets["Additions"].columns) == ["new"]


# --------------------------------------------------------------------------
# Loader
# --------------------------------------------------------------------------

def test_header_detection_skips_title_rows(tmp_path):
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "CAN"
    sheet.append(["Amounts in CAD$"])
    sheet.append([])
    sheet.append([])
    sheet.append(["Unit Number", "Serial Number", "Facility", "Model Year"])
    for i in range(20):
        sheet.append([f"U{i}", f"S{i}", "LACH", 2020 + (i % 5)])
    path = tmp_path / "titled.xlsx"
    workbook.save(path)

    info = detect_header_row(path, "CAN")
    assert info.header_row == 4
    assert info.columns[:2] == ["Unit Number", "Serial Number"]

    frame = read_sheet(path, "CAN")
    assert len(frame) == 20
    assert list(frame.columns) == ["Unit Number", "Serial Number", "Facility", "Model Year"]
