"""Leasing asset-list comparison app by Suraj Shah.

Run with:  streamlit run app.py
"""

from __future__ import annotations

import re
import tempfile
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

from src.calc import (
    BUILDER_OPERATORS,
    MULTI_OPERATORS,
    FUNCTIONS,
    Operation,
    apply_operations,
    build_binary_formula,
    build_multi_formula,
    operand_expression,
)
from src.cleaning import CleaningOptions, clean_dataframe, suggest_key_columns
from src.diffing import DiffOptions, column_drift, compare_sheets
from src.exporter import SheetPayload, append_sheets, write_new_workbook
from src.loader import formula_leak_ratio, profile_workbook, read_sheet
from src.naming import build_column_map, pair_sheets

st.set_page_config(
    page_title="Fleet Asset List Comparison",
    page_icon="📊",
    layout="wide",
)

DATE_IN_NAME = re.compile(r"(\d{1,2})[-_.](\d{1,2})[-_.](\d{2,4})")


# ---------------------------------------------------------------------------
# Session helpers
# ---------------------------------------------------------------------------

def session_dir() -> Path:
    if "workdir" not in st.session_state:
        st.session_state.workdir = tempfile.mkdtemp(prefix="ten_assets_")
    return Path(st.session_state.workdir)


def init_state() -> None:
    defaults = {
        "paths": {},
        "profiles": {},
        "cleaned": {},          # label -> DataFrame
        "clean_reports": {},    # label -> CleaningReport
        "diffs": {},            # new sheet -> DiffResult
        "operations": [],       # list of dicts
        "calc_frames": {},      # label -> DataFrame with calculated columns
        "export_path": None,
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


init_state()


def date_from_filename(name: str) -> Optional[datetime]:
    match = DATE_IN_NAME.search(Path(name).stem)
    if not match:
        return None
    month, day, year = (int(g) for g in match.groups())
    if year < 100:
        year += 2000
    try:
        return datetime(year, month, day)
    except ValueError:
        return None


def save_upload(uploaded, label: str) -> Path:
    path = session_dir() / f"{label}_{uploaded.name}"
    if not path.exists() or path.stat().st_size != uploaded.size:
        with open(path, "wb") as handle:
            handle.write(uploaded.getbuffer())
    return path


# ---------------------------------------------------------------------------
# Cached heavy operations
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner="Inspecting workbook…")
def cached_profile(path: str, size: int) -> Dict[str, Tuple[int, int, List[str]]]:
    profile = profile_workbook(path)
    return {
        name: (info.header_row, info.n_columns, info.columns)
        for name, info in profile.items()
    }


@st.cache_data(show_spinner="Reading sheet…")
def cached_read(path: str, size: int, sheet: str, header_row: int) -> pd.DataFrame:
    return read_sheet(path, sheet, header_row=header_row)


@st.cache_data(show_spinner="Cleaning…")
def cached_clean(path: str, size: int, sheet: str, header_row: int, options_key: tuple):
    """Read and clean in one cached step.

    Keyed on the file rather than the DataFrame -- hashing a 58k x 122 frame on
    every rerun costs more than the cleaning itself.
    """
    raw = cached_read(path, size, sheet, header_row)
    cleaned, report = clean_dataframe(raw, CleaningOptions(**dict(options_key)))
    return cleaned, report, formula_leak_ratio(cleaned)


# ---------------------------------------------------------------------------
# Sidebar: file intake
# ---------------------------------------------------------------------------

st.sidebar.title("📊 Asset List Comparison- by Suraj Shah")
st.sidebar.caption("Compare two dated asset lists, then build calculated columns.")

uploads = st.sidebar.file_uploader(
    "Upload both asset lists (.xlsx)",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    help="Upload exactly two files. The newer one is detected from the date in the filename.",
)

st.sidebar.divider()

if not uploads:
    st.title("Asset List Comparison - by Suraj Shah")
    st.info(
        "Upload two dated asset-list workbooks in the sidebar to begin. "
        "The app figures out which is newer, lines the sheets up, cleans both, "
        "and reports what was added and removed."
    )
    st.markdown(
        """
        **What this app does**

        1. **Cleans** both files — trims padding and tabs, converts numbers and dates
           stored as text, and blanks placeholder values such as `NULL`, `--` and `1900-01-01`.
        2. **New Additions** — assets present in the newer file but not the older one.
        3. **Deletions** — assets present in the older file but gone from the newer one.
        4. **Calculated columns** — combine any columns with any operation and name the result.
        5. **Exports** everything as new tabs inside the newer workbook, leaving its
           original sheets untouched.
        """
    )
    st.stop()

if len(uploads) != 2:
    st.sidebar.error(f"Please upload exactly two files (got {len(uploads)}).")
    st.stop()

# Decide which file is newer.
dated = [(date_from_filename(u.name), u) for u in uploads]
if all(d is not None for d, _ in dated):
    dated.sort(key=lambda pair: pair[0])
    default_new = dated[-1][1].name
else:
    default_new = uploads[0].name

newer_name = st.sidebar.radio(
    "Which file is the newer one?",
    [u.name for u in uploads],
    index=[u.name for u in uploads].index(default_new),
    help="Detected from the date in the filename; override if it guessed wrong.",
)

new_upload = next(u for u in uploads if u.name == newer_name)
old_upload = next(u for u in uploads if u.name != newer_name)

new_path = save_upload(new_upload, "new")
old_path = save_upload(old_upload, "old")
st.session_state.paths = {"new": str(new_path), "old": str(old_path)}

st.sidebar.success(f"**Newer:** {new_upload.name}")
st.sidebar.info(f"**Older:** {old_upload.name}")

new_profile = cached_profile(str(new_path), new_upload.size)
old_profile = cached_profile(str(old_path), old_upload.size)

# ---------------------------------------------------------------------------
# Main tabs
# ---------------------------------------------------------------------------

tab_setup, tab_diff, tab_calc, tab_export = st.tabs(
    ["1 · Clean", "2 · Additions & Deletions", "3 · Calculated columns", "4 · Export"]
)


# =========================== 1. CLEAN ======================================
with tab_setup:
    st.header("Sheet pairing and data cleaning")

    suggestions = pair_sheets(list(new_profile), list(old_profile))
    # Summary and pivot tabs are narrow and aggregated -- never the asset register.
    data_sheets = [
        name
        for name, (_, n_cols, _) in new_profile.items()
        if n_cols >= 10 and not re.search(r"summary|pivot|chart|notes", name, re.I)
    ]

    st.subheader("Which sheets to compare")
    selected_sheets = st.multiselect(
        "Sheets in the newer file",
        options=list(new_profile),
        default=[s for s in data_sheets if s in dict(suggestions) and dict(suggestions)[s]],
        help="Summary or pivot sheets are usually left out.",
    )

    pairing: Dict[str, Optional[str]] = {}
    if selected_sheets:
        columns = st.columns(min(len(selected_sheets), 3))
        suggested = dict(suggestions)
        for index, sheet in enumerate(selected_sheets):
            with columns[index % len(columns)]:
                options = ["(none)"] + list(old_profile)
                guess = suggested.get(sheet)
                pick = st.selectbox(
                    f"Older-file match for **{sheet}**",
                    options,
                    index=options.index(guess) if guess in options else 0,
                    key=f"pair_{sheet}",
                )
                pairing[sheet] = None if pick == "(none)" else pick
                head_new = new_profile[sheet][0]
                head_old = old_profile[pick][0] if pick != "(none)" else "-"
                st.caption(
                    f"Header rows — newer: {head_new}, older: {head_old} "
                    f"· {new_profile[sheet][1]} columns"
                )

    with st.expander("Header row overrides", expanded=False):
        st.caption(
            "Header rows are detected automatically. Override only if a preview below "
            "looks wrong."
        )
        overrides: Dict[str, Dict[str, int]] = {}
        for sheet in selected_sheets:
            left, right = st.columns(2)
            with left:
                overrides.setdefault(sheet, {})["new"] = st.number_input(
                    f"{sheet} (newer) header row",
                    min_value=1, max_value=50,
                    value=new_profile[sheet][0], key=f"hr_new_{sheet}",
                )
            with right:
                partner = pairing.get(sheet)
                overrides[sheet]["old"] = st.number_input(
                    f"{partner or '—'} (older) header row",
                    min_value=1, max_value=50,
                    value=old_profile[partner][0] if partner else 1,
                    key=f"hr_old_{sheet}",
                    disabled=partner is None,
                )

    st.subheader("Cleaning steps")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        trim = st.checkbox("Trim leading/trailing spaces & tabs", value=True)
        collapse = st.checkbox("Collapse repeated inner spaces", value=True)
        nulls = st.checkbox("Blank out NULL, --, N/A, #N/A", value=True)
    with col_b:
        numeric = st.checkbox("Convert numbers stored as text", value=True)
        dates = st.checkbox("Convert dates stored as text", value=True)
        placeholder = st.checkbox("Blank placeholder dates (1900-01-01)", value=True)
    with col_c:
        blank_rows = st.checkbox("Drop fully blank rows", value=True)
        blank_cols = st.checkbox("Drop fully blank columns", value=True)
        dupes = st.checkbox("Drop exact duplicate rows", value=False)

    options = CleaningOptions(
        trim_whitespace=trim,
        collapse_inner_spaces=collapse,
        normalise_nulls=nulls,
        drop_blank_rows=blank_rows,
        drop_blank_columns=blank_cols,
        convert_numeric=numeric,
        convert_dates=dates,
        blank_placeholder_dates=placeholder,
        drop_duplicate_rows=dupes,
    )
    options_key = tuple(sorted(asdict(options).items()))

    if st.button("Run cleaning", type="primary", disabled=not selected_sheets):
        st.session_state.cleaned = {}
        st.session_state.clean_reports = {}
        st.session_state.diffs = {}
        st.session_state.calc_frames = {}

        progress = st.progress(0.0, text="Starting…")
        jobs: List[Tuple[str, str, str, int, int]] = []
        for sheet in selected_sheets:
            jobs.append(("new", sheet, str(new_path), new_upload.size, overrides[sheet]["new"]))
            partner = pairing.get(sheet)
            if partner:
                jobs.append(
                    ("old", partner, str(old_path), old_upload.size, overrides[sheet]["old"])
                )

        leaks: Dict[str, float] = {}
        for index, (side, sheet, path, size, header_row) in enumerate(jobs):
            label = f"{side}:{sheet}"
            progress.progress(index / len(jobs), text=f"Reading and cleaning {sheet}…")
            cleaned, report, leak = cached_clean(path, size, sheet, header_row, options_key)
            st.session_state.cleaned[label] = cleaned
            st.session_state.clean_reports[label] = report
            leaks[label] = leak
        st.session_state.formula_leaks = leaks
        progress.progress(1.0, text="Done")
        st.session_state.pairing = pairing
        st.success(f"Cleaned {len(jobs)} sheet(s).")

    if st.session_state.clean_reports:
        st.subheader("Cleaning report")
        for label, report in st.session_state.clean_reports.items():
            side, sheet = label.split(":", 1)
            side_name = "newer file" if side == "new" else "older file"
            with st.expander(f"**{sheet}** ({side_name})", expanded=side == "new"):
                summary = pd.DataFrame(report.as_records())
                left, right = st.columns([1, 2])
                with left:
                    st.dataframe(summary, hide_index=True, width="stretch")
                with right:
                    if report.converted_to_numeric:
                        st.markdown("**Turned into numbers:** " + ", ".join(
                            f"`{c}`" for c in report.converted_to_numeric))
                    if report.converted_to_datetime:
                        st.markdown("**Turned into dates:** " + ", ".join(
                            f"`{c}`" for c in report.converted_to_datetime))
                    if report.columns_dropped:
                        st.markdown("**Empty columns dropped:** " + ", ".join(
                            f"`{c}`" for c in report.columns_dropped))
                    if report.placeholder_dates_blanked:
                        st.markdown("**Placeholder dates blanked:** " + ", ".join(
                            f"`{c}` ({n:,})"
                            for c, n in report.placeholder_dates_blanked.items()))

                frame = st.session_state.cleaned[label]
                leak = st.session_state.get("formula_leaks", {}).get(label, 0.0)
                if leak > 0.001:
                    st.warning(
                        f"{leak:.1%} of text cells still hold raw formulas. The source "
                        "workbook was saved without cached results — open it in Excel, "
                        "recalculate and save before relying on those columns."
                    )
                st.dataframe(frame.head(200), width="stretch", height=260)


# ==================== 2. ADDITIONS & DELETIONS =============================
with tab_diff:
    st.header("New additions and deletions")

    if not st.session_state.cleaned:
        st.info("Run cleaning on the first tab to unlock this step.")
    else:
        pairing = st.session_state.get("pairing", {})
        paired = [s for s, p in pairing.items() if p]
        if not paired:
            st.warning("No sheet in the newer file is matched to one in the older file.")
        for sheet in paired:
            partner = pairing[sheet]
            new_frame = st.session_state.cleaned[f"new:{sheet}"]
            old_frame = st.session_state.cleaned[f"old:{partner}"]

            st.subheader(f"{sheet}  ⟷  {partner}")
            column_pairs = build_column_map(list(new_frame.columns), list(old_frame.columns))
            defaults = suggest_key_columns(list(new_frame.columns))[:1]

            left, right = st.columns([2, 1])
            with left:
                keys = st.multiselect(
                    "Match assets on",
                    options=list(new_frame.columns),
                    default=defaults,
                    key=f"keys_{sheet}",
                    help="Serial number is usually the most reliable identifier. "
                         "Pick more than one to match on a combination.",
                )
            with right:
                strip_zeros = st.checkbox(
                    "Ignore leading zeros", value=False, key=f"zeros_{sheet}"
                )
                st.caption("Whitespace and letter case are always ignored.")

            unmatched = [k for k in keys if column_pairs.get(k) is None]
            if unmatched:
                st.error(
                    "No equivalent in the older file for: "
                    + ", ".join(f"`{k}`" for k in unmatched)
                )
            elif keys and st.button("Compare", key=f"cmp_{sheet}", type="primary"):
                result = compare_sheets(
                    new_frame,
                    old_frame,
                    DiffOptions(
                        key_columns=keys,
                        old_key_columns=[column_pairs[k] for k in keys],
                        strip_leading_zeros=strip_zeros,
                    ),
                )
                st.session_state.diffs[sheet] = result

            result = st.session_state.diffs.get(sheet)
            if result is not None:
                metrics = st.columns(5)
                metrics[0].metric("Rows — newer", f"{result.n_new_rows:,}")
                metrics[1].metric("Rows — older", f"{result.n_old_rows:,}")
                metrics[2].metric("In both", f"{result.n_common:,}")
                metrics[3].metric("New additions", f"{len(result.additions):,}")
                metrics[4].metric("Deletions", f"{len(result.deletions):,}")

                if result.n_duplicate_keys_new or result.n_duplicate_keys_old:
                    st.warning(
                        f"Duplicate keys found — {result.n_duplicate_keys_new} in the newer "
                        f"file, {result.n_duplicate_keys_old} in the older. "
                        + (f"Examples: {', '.join(result.duplicate_examples)}"
                           if result.duplicate_examples else "")
                    )
                if result.n_missing_keys_new or result.n_missing_keys_old:
                    st.warning(
                        f"Rows with a blank key were skipped — "
                        f"{result.n_missing_keys_new} newer, {result.n_missing_keys_old} older."
                    )

                view_add, view_del, view_drift = st.tabs(
                    [
                        f"New additions ({len(result.additions):,})",
                        f"Deletions ({len(result.deletions):,})",
                        "Column changes",
                    ]
                )
                with view_add:
                    st.dataframe(result.additions.head(500), width="stretch", height=320)
                    st.caption("Showing the first 500 rows; the export contains all of them.")
                    st.download_button(
                        "Download additions as CSV",
                        result.additions.to_csv(index=False).encode("utf-8"),
                        file_name=f"{sheet}_new_additions.csv",
                        mime="text/csv",
                        key=f"dl_add_{sheet}",
                    )
                with view_del:
                    st.dataframe(result.deletions.head(500), width="stretch", height=320)
                    st.download_button(
                        "Download deletions as CSV",
                        result.deletions.to_csv(index=False).encode("utf-8"),
                        file_name=f"{sheet}_deletions.csv",
                        mime="text/csv",
                        key=f"dl_del_{sheet}",
                    )
                with view_drift:
                    drift = column_drift(list(new_frame.columns), list(old_frame.columns))
                    if drift.empty:
                        st.success("Both files carry exactly the same columns.")
                    else:
                        st.dataframe(drift, hide_index=True, width="stretch", height=320)
            st.divider()


# ======================= 3. CALCULATED COLUMNS =============================
with tab_calc:
    st.header("Build calculated columns")

    available: Dict[str, pd.DataFrame] = {}
    for label, frame in st.session_state.cleaned.items():
        side, sheet = label.split(":", 1)
        if side == "new":
            available[f"{sheet} (full sheet)"] = frame
    for sheet, result in st.session_state.diffs.items():
        available[f"{sheet} — New Additions"] = result.additions
        available[f"{sheet} — Deletions"] = result.deletions

    if not available:
        st.info("Run cleaning on the first tab to unlock this step.")
    else:
        target_label = st.selectbox("Apply calculations to", list(available))
        target = available[target_label]
        columns = list(target.columns)
        numeric_columns = [
            c for c in columns if pd.api.types.is_numeric_dtype(target[c])
        ]

        st.caption(
            f"{len(target):,} rows · {len(columns)} columns · "
            f"{len(numeric_columns)} of them numeric"
        )

        st.subheader("Add an operation")
        mode = st.radio(
            "Operation type",
            ["Two operands", "Several columns", "Custom formula"],
            horizontal=True,
        )

        formula = ""
        if mode == "Two operands":
            row = st.columns([3, 2, 3])
            with row[0]:
                left_kind = st.radio("Left side", ["Column", "Number"], horizontal=True,
                                     key="lk")
                left_value = (
                    st.selectbox("Left column", columns, key="lv")
                    if left_kind == "Column"
                    else st.text_input("Left number", value="1", key="lvn")
                )
            with row[1]:
                operator = st.selectbox("Operation", list(BUILDER_OPERATORS), key="op")
            with row[2]:
                right_kind = st.radio("Right side", ["Column", "Number"], horizontal=True,
                                      key="rk")
                right_value = (
                    st.selectbox(
                        "Right column",
                        columns,
                        index=min(1, len(columns) - 1),
                        key="rv",
                    )
                    if right_kind == "Column"
                    else st.text_input("Right number", value="1", key="rvn")
                )
            formula = build_binary_formula(
                operator,
                operand_expression(left_kind, left_value),
                operand_expression(right_kind, right_value),
            )

        elif mode == "Several columns":
            operator = st.selectbox("Operation", list(MULTI_OPERATORS), key="mop")
            picked = st.multiselect(
                "Columns", columns, default=numeric_columns[:2], key="mcols"
            )
            if picked:
                formula = build_multi_formula(operator, picked)

        else:
            formula = st.text_area(
                "Formula",
                value="([$ OLV] - [NBV_Current]) / [NBV_Current] * 100",
                height=90,
                help="Reference columns in square brackets.",
            )
            with st.expander("Formula reference"):
                st.markdown(
                    "**Columns** — write them in square brackets: `[Monthly Lease Rate]`\n\n"
                    "**Operators** — `+  -  *  /  //  %  **` and comparisons "
                    "`>  >=  <  <=  ==  !=`\n\n"
                    "**Row-by-row functions** — `SUM`, `AVG`, `MIN`, `MAX`, `PRODUCT`, "
                    "`ABS`, `ROUND(x, n)`, `SQRT`, `LN`, `LOG10`, `FLOOR`, `CEIL`, "
                    "`IF(test, a, b)`, `COALESCE`, `CONCAT`, `ISBLANK`, `COUNTVALUES`, "
                    "`DAYS(a, b)`, `YEAR`, `MONTH`\n\n"
                    "**Whole-column totals** — `COLSUM`, `COLAVG`, `COLMIN`, `COLMAX`, "
                    "`COLMEDIAN`, `COLCOUNT`\n\n"
                    "**Examples**\n"
                    "```\n"
                    "[MonthlyLeaseRate] + [MonthlyRentRate]\n"
                    "ROUND([$ OLV] / COLSUM([$ OLV]) * 100, 2)\n"
                    "IF([ModelYear] >= 2020, [$ OLV] * 1.1, [$ OLV])\n"
                    "DAYS([ExpectedReturnDate], [ContractDate])\n"
                    "```"
                )

        name_col, round_col, add_col = st.columns([3, 1, 1])
        with name_col:
            output_name = st.text_input("Name the new column", value="", key="outname")
        with round_col:
            rounding = st.selectbox(
                "Round to", ["No rounding", "0", "1", "2", "3", "4"], index=0
            )
        with add_col:
            st.write("")
            st.write("")
            add_clicked = st.button("Add operation", type="primary")

        if formula:
            st.code(formula, language="text")

        if add_clicked:
            if not output_name.strip():
                st.error("Give the new column a name.")
            elif not formula.strip():
                st.error("Build a formula first.")
            else:
                st.session_state.operations.append(
                    {
                        "output_name": output_name.strip(),
                        "formula": formula,
                        "round_to": None if rounding == "No rounding" else int(rounding),
                        "target": target_label,
                    }
                )
                st.rerun()

        st.subheader("Queued operations")
        queued = [o for o in st.session_state.operations if o["target"] == target_label]
        if not queued:
            st.caption(
                "Nothing queued yet. Operations run in order, and each new column "
                "can be used by the ones after it."
            )
        else:
            for index, op in enumerate(queued):
                row = st.columns([2, 5, 1, 1])
                row[0].markdown(f"**{op['output_name']}**")
                row[1].code(op["formula"], language="text")
                row[2].caption(
                    "round " + (str(op["round_to"]) if op["round_to"] is not None else "—")
                )
                if row[3].button("Remove", key=f"rm_{target_label}_{index}"):
                    st.session_state.operations.remove(op)
                    st.rerun()

            action = st.columns([1, 1, 4])
            if action[0].button("Apply all", type="primary"):
                result = apply_operations(
                    target,
                    [
                        Operation(
                            output_name=o["output_name"],
                            formula=o["formula"],
                            round_to=o["round_to"],
                        )
                        for o in queued
                    ],
                )
                st.session_state.calc_frames[target_label] = result.frame
                if result.applied:
                    st.success("Added: " + ", ".join(f"`{c}`" for c in result.applied))
                for name, message in result.errors:
                    st.error(f"**{name}** — {message}")
            if action[1].button("Clear all"):
                st.session_state.operations = [
                    o for o in st.session_state.operations if o["target"] != target_label
                ]
                st.rerun()

        calculated = st.session_state.calc_frames.get(target_label)
        if calculated is not None:
            new_columns = [c for c in calculated.columns if c not in target.columns]
            st.subheader("Preview")
            preview_columns = (
                [c for c in calculated.columns if c in columns][:4] + new_columns
            )
            st.dataframe(
                calculated[preview_columns].head(200), width="stretch", height=320
            )
            if new_columns:
                stats = calculated[new_columns].describe(include="all").T
                st.caption("Summary of the calculated columns")
                st.dataframe(stats, width="stretch")


# ============================== 4. EXPORT ==================================
with tab_export:
    st.header("Export")

    payload_options: Dict[str, pd.DataFrame] = {}
    for sheet, result in st.session_state.diffs.items():
        payload_options[f"{sheet} — New Additions"] = result.additions
        payload_options[f"{sheet} — Deletions"] = result.deletions
    for label, frame in st.session_state.calc_frames.items():
        payload_options[f"{label} + calculations"] = frame

    if not payload_options:
        st.info("Compare at least one pair of sheets to have something to export.")
    else:
        chosen = st.multiselect(
            "Tabs to write",
            list(payload_options),
            default=[k for k in payload_options if "calculations" not in k],
        )

        st.subheader("Tab names")
        names: Dict[str, str] = {}
        for key in chosen:
            default_name = key.split("—")[-1].strip()
            if "calculations" in key:
                default_name = key.replace(" + calculations", "").split("(")[0].strip()
                default_name = f"{default_name} Calc"[:31]
            names[key] = st.text_input(
                f"Sheet name for *{key}*", value=default_name[:31], key=f"name_{key}",
                max_chars=31,
            )

        mode = st.radio(
            "Where to write them",
            [
                "Add tabs to the newer workbook (keeps every original sheet)",
                "Create a small new workbook with just these tabs",
            ],
        )
        if mode.startswith("Add tabs"):
            st.caption(
                "The original sheets are copied across untouched — formatting, formulas "
                "and pivots all survive. A tab whose name already exists is replaced."
            )
        else:
            st.caption("Much smaller and faster to download, but the source sheets are not included.")

        if st.button("Build the file", type="primary", disabled=not chosen):
            payloads = [
                SheetPayload(name=names[key], frame=payload_options[key]) for key in chosen
            ]
            stamp = datetime.now().strftime("%Y%m%d_%H%M")
            source = Path(st.session_state.paths["new"])
            with st.spinner("Writing workbook…"):
                if mode.startswith("Add tabs"):
                    out = session_dir() / f"{source.stem}_compared_{stamp}.xlsx"
                    written = append_sheets(source, out, payloads)
                else:
                    out = session_dir() / f"asset_comparison_{stamp}.xlsx"
                    written = write_new_workbook(out, payloads)
            st.session_state.export_path = str(out)
            st.success("Wrote: " + ", ".join(f"`{n}`" for n in written))

        export_path = st.session_state.get("export_path")
        if export_path and Path(export_path).exists():
            size_mb = Path(export_path).stat().st_size / 1e6
            with open(export_path, "rb") as handle:
                st.download_button(
                    f"Download {Path(export_path).name}  ({size_mb:.1f} MB)",
                    handle.read(),
                    file_name=Path(export_path).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
