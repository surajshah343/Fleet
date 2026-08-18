"""User-defined calculated columns.

Two ways to describe a calculation, both compiled to the same expression and
evaluated by the same sandbox:

* **Builder** -- pick a left operand, an operator and a right operand.
  ``Monthly Lease Rate`` + ``Monthly Rent Rate`` -> ``Total Rate``.
* **Formula** -- type an Excel-ish expression referencing columns in square
  brackets: ``([$ OLV] - [NBV_Current]) / [NBV_Current] * 100``.

Evaluation walks a parsed syntax tree and refuses anything not on the
whitelist, so a formula can never reach the filesystem, imports or attributes.
``eval`` on raw user text is never used.
"""

from __future__ import annotations

import ast
import math
import re
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd

COLUMN_TOKEN = re.compile(r"\[([^\[\]]+)\]")
_PLACEHOLDER = "__col_{}__"


# --------------------------------------------------------------------------
# Operators exposed by the point-and-click builder
# --------------------------------------------------------------------------

BUILDER_OPERATORS: Dict[str, str] = {
    "Add (+)": "{left} + {right}",
    "Subtract (-)": "{left} - {right}",
    "Multiply (*)": "{left} * {right}",
    "Divide (/)": "{left} / {right}",
    "Power (^)": "{left} ** {right}",
    "Remainder (mod)": "{left} % {right}",
    "Minimum": "MIN({left}, {right})",
    "Maximum": "MAX({left}, {right})",
    "Average": "(({left}) + ({right})) / 2",
    "Percent of": "{left} / {right} * 100",
    "Percent change": "({right} - {left}) / {left} * 100",
    "Difference (absolute)": "ABS({left} - {right})",
    "Concatenate (text)": "CONCAT({left}, {right})",
}

MULTI_OPERATORS: Dict[str, str] = {
    "Sum of columns": "SUM({args})",
    "Average of columns": "AVG({args})",
    "Minimum of columns": "MIN({args})",
    "Maximum of columns": "MAX({args})",
    "Multiply columns": "PRODUCT({args})",
    "Count filled values": "COUNTVALUES({args})",
}


# --------------------------------------------------------------------------
# Function library available inside formulas
# --------------------------------------------------------------------------

def _to_number(value):
    """Coerce a Series/scalar to numbers, turning junk into NaN."""
    if isinstance(value, pd.Series):
        if pd.api.types.is_numeric_dtype(value):
            return value
        if pd.api.types.is_datetime64_any_dtype(value):
            return value
        text = value.astype("string").str.replace(r"[,$€£%\s]", "", regex=True)
        negative = text.str.match(r"^\(.*\)$", na=False)
        text = text.str.replace(r"[()]", "", regex=True)
        numbers = pd.to_numeric(text, errors="coerce")
        return numbers.where(~negative, -numbers.abs())
    return value


def _align(args: Sequence) -> List:
    return [_to_number(a) for a in args]


def _rowwise(func: Callable, args: Sequence):
    """Apply a numpy reduction across columns, row by row."""
    numeric = _align(args)
    series = [a for a in numeric if isinstance(a, pd.Series)]
    if not series:
        return func(np.array([[float(a) for a in numeric]]), axis=1)[0]
    frame = pd.concat(
        [
            a if isinstance(a, pd.Series) else pd.Series(a, index=series[0].index)
            for a in numeric
        ],
        axis=1,
    )
    return pd.Series(func(frame.to_numpy(dtype="float64"), axis=1), index=frame.index)


def _if(condition, when_true, when_false):
    if isinstance(condition, pd.Series):
        return pd.Series(np.where(condition.fillna(False), when_true, when_false),
                         index=condition.index)
    return when_true if condition else when_false


def _coalesce(*args):
    result = args[0]
    for candidate in args[1:]:
        if isinstance(result, pd.Series):
            result = result.fillna(candidate)
        elif result is None or (isinstance(result, float) and math.isnan(result)):
            result = candidate
    return result


def _concat(*args):
    parts = []
    for arg in args:
        if isinstance(arg, pd.Series):
            parts.append(arg.astype("string").fillna(""))
        else:
            parts.append(str(arg))
    series = [p for p in parts if isinstance(p, pd.Series)]
    if not series:
        return "".join(parts)
    index = series[0].index
    out = pd.Series("", index=index, dtype="string")
    for part in parts:
        out = out.str.cat(
            part if isinstance(part, pd.Series) else pd.Series(part, index=index),
            na_rep="",
        )
    return out


def _count_values(*args):
    counts = None
    for arg in args:
        filled = arg.notna().astype(int) if isinstance(arg, pd.Series) else int(arg is not None)
        counts = filled if counts is None else counts + filled
    return counts


def _column_agg(func: Callable):
    """Whole-column aggregate that broadcasts back across every row."""

    def wrapper(value):
        numeric = _to_number(value)
        if isinstance(numeric, pd.Series):
            return func(numeric)
        return numeric

    return wrapper


FUNCTIONS: Dict[str, Callable] = {
    "ABS": lambda x: _to_number(x).abs() if isinstance(x, pd.Series) else abs(x),
    "ROUND": lambda x, n=0: (
        _to_number(x).round(int(n)) if isinstance(x, pd.Series) else round(x, int(n))
    ),
    "FLOOR": lambda x: np.floor(_to_number(x)),
    "CEIL": lambda x: np.ceil(_to_number(x)),
    "SQRT": lambda x: np.sqrt(_to_number(x)),
    "EXP": lambda x: np.exp(_to_number(x)),
    "LN": lambda x: np.log(_to_number(x)),
    "LOG10": lambda x: np.log10(_to_number(x)),
    "SUM": lambda *a: _rowwise(np.nansum, a),
    "AVG": lambda *a: _rowwise(np.nanmean, a),
    "MEAN": lambda *a: _rowwise(np.nanmean, a),
    "MIN": lambda *a: _rowwise(np.nanmin, a),
    "MAX": lambda *a: _rowwise(np.nanmax, a),
    "PRODUCT": lambda *a: _rowwise(np.nanprod, a),
    "COUNTVALUES": _count_values,
    "IF": _if,
    "COALESCE": _coalesce,
    "CONCAT": _concat,
    "ISBLANK": lambda x: x.isna() if isinstance(x, pd.Series) else x is None,
    "NUMBER": _to_number,
    # Whole-column aggregates, useful for share-of-total columns.
    "COLSUM": _column_agg(lambda s: s.sum(skipna=True)),
    "COLAVG": _column_agg(lambda s: s.mean(skipna=True)),
    "COLMIN": _column_agg(lambda s: s.min(skipna=True)),
    "COLMAX": _column_agg(lambda s: s.max(skipna=True)),
    "COLMEDIAN": _column_agg(lambda s: s.median(skipna=True)),
    "COLCOUNT": _column_agg(lambda s: s.notna().sum()),
    "DAYS": lambda a, b: (pd.to_datetime(a) - pd.to_datetime(b)).dt.days,
    "YEAR": lambda x: pd.to_datetime(x).dt.year,
    "MONTH": lambda x: pd.to_datetime(x).dt.month,
}

CONSTANTS: Dict[str, float] = {"PI": math.pi, "E": math.e, "TRUE": True, "FALSE": False}

_ALLOWED_BINOPS = {
    ast.Add: lambda a, b: a + b,
    ast.Sub: lambda a, b: a - b,
    ast.Mult: lambda a, b: a * b,
    ast.Div: lambda a, b: a / b,
    ast.FloorDiv: lambda a, b: a // b,
    ast.Mod: lambda a, b: a % b,
    ast.Pow: lambda a, b: a ** b,
}

_ALLOWED_COMPARE = {
    ast.Gt: lambda a, b: a > b,
    ast.GtE: lambda a, b: a >= b,
    ast.Lt: lambda a, b: a < b,
    ast.LtE: lambda a, b: a <= b,
    ast.Eq: lambda a, b: a == b,
    ast.NotEq: lambda a, b: a != b,
}


class FormulaError(ValueError):
    """Raised when a formula is malformed or uses something not allowed."""


# --------------------------------------------------------------------------
# Parsing and evaluation
# --------------------------------------------------------------------------

def encode_columns(formula: str, columns: Sequence[str]) -> Tuple[str, Dict[str, str]]:
    """Swap ``[Column Name]`` references for safe identifiers.

    Bare column names are also honoured when they happen to be valid Python
    identifiers and are not shadowed by a function name.
    """
    mapping: Dict[str, str] = {}
    by_lower = {c.lower(): c for c in columns}

    def replace(match: "re.Match[str]") -> str:
        raw = match.group(1).strip()
        column = by_lower.get(raw.lower())
        if column is None:
            raise FormulaError(f"Unknown column: [{raw}]")
        placeholder = _PLACEHOLDER.format(len(mapping))
        mapping[placeholder] = column
        return placeholder

    encoded = COLUMN_TOKEN.sub(replace, formula)

    for column in sorted(columns, key=len, reverse=True):
        if column.isidentifier() and column.upper() not in FUNCTIONS:
            pattern = rf"(?<![\w\[]){re.escape(column)}(?![\w\]])"
            if re.search(pattern, encoded):
                placeholder = _PLACEHOLDER.format(len(mapping))
                mapping[placeholder] = column
                encoded = re.sub(pattern, placeholder, encoded)
    return encoded, mapping


def _evaluate_node(node: ast.AST, scope: Dict[str, object]):
    if isinstance(node, ast.Expression):
        return _evaluate_node(node.body, scope)
    if isinstance(node, ast.Constant):
        if isinstance(node.value, (int, float, bool, str)):
            return node.value
        raise FormulaError(f"Unsupported literal: {node.value!r}")
    if isinstance(node, ast.Name):
        if node.id in scope:
            return scope[node.id]
        if node.id.upper() in CONSTANTS:
            return CONSTANTS[node.id.upper()]
        raise FormulaError(f"Unknown name: {node.id}")
    if isinstance(node, ast.BinOp):
        handler = _ALLOWED_BINOPS.get(type(node.op))
        if handler is None:
            raise FormulaError(f"Operator not allowed: {type(node.op).__name__}")
        # Arithmetic always works on numbers.  A column of "1,200.50" strings
        # would otherwise be concatenated by `+` rather than added.  Use
        # CONCAT() when text joining is what you actually want.
        left = _to_number(_evaluate_node(node.left, scope))
        right = _to_number(_evaluate_node(node.right, scope))
        return handler(left, right)
    if isinstance(node, ast.UnaryOp):
        operand = _evaluate_node(node.operand, scope)
        if isinstance(node.op, ast.USub):
            return -operand
        if isinstance(node.op, ast.UAdd):
            return operand
        if isinstance(node.op, ast.Not):
            return ~operand if isinstance(operand, pd.Series) else (not operand)
        raise FormulaError("Unary operator not allowed")
    if isinstance(node, ast.Compare):
        if len(node.ops) != 1:
            raise FormulaError("Chained comparisons are not supported")
        handler = _ALLOWED_COMPARE.get(type(node.ops[0]))
        if handler is None:
            raise FormulaError("Comparison not allowed")
        left = _evaluate_node(node.left, scope)
        right = _evaluate_node(node.comparators[0], scope)
        # Comparing a text column against a number means the column holds
        # numbers stored as text; comparing against text stays textual.
        if isinstance(right, (int, float)) and not isinstance(right, bool):
            left = _to_number(left)
        if isinstance(left, (int, float)) and not isinstance(left, bool):
            right = _to_number(right)
        return handler(left, right)
    if isinstance(node, ast.BoolOp):
        values = [_evaluate_node(v, scope) for v in node.values]
        result = values[0]
        for value in values[1:]:
            result = (result & value) if isinstance(node.op, ast.And) else (result | value)
        return result
    if isinstance(node, ast.Call):
        if not isinstance(node.func, ast.Name):
            raise FormulaError("Only plain function calls are allowed")
        name = node.func.id.upper()
        function = FUNCTIONS.get(name)
        if function is None:
            raise FormulaError(f"Unknown function: {node.func.id}")
        if node.keywords:
            raise FormulaError("Keyword arguments are not supported")
        return function(*[_evaluate_node(a, scope) for a in node.args])
    raise FormulaError(f"Expression element not allowed: {type(node).__name__}")


def evaluate_formula(frame: pd.DataFrame, formula: str) -> pd.Series:
    """Evaluate a formula against ``frame`` and return the resulting column."""
    if not formula or not formula.strip():
        raise FormulaError("Formula is empty")

    encoded, mapping = encode_columns(formula, list(frame.columns))
    try:
        tree = ast.parse(encoded, mode="eval")
    except SyntaxError as exc:
        raise FormulaError(f"Could not parse formula: {exc.msg}") from exc

    scope: Dict[str, object] = {
        placeholder: frame[column] for placeholder, column in mapping.items()
    }
    result = _evaluate_node(tree, scope)

    if not isinstance(result, pd.Series):
        result = pd.Series([result] * len(frame), index=frame.index)
    if pd.api.types.is_numeric_dtype(result):
        result = result.replace([np.inf, -np.inf], np.nan)
    return result


# --------------------------------------------------------------------------
# Operation objects used by the app
# --------------------------------------------------------------------------

@dataclass
class Operation:
    """One calculated column."""

    output_name: str
    formula: str
    description: str = ""
    round_to: Optional[int] = None

    def compiled(self) -> str:
        if self.round_to is None:
            return self.formula
        return f"ROUND({self.formula}, {int(self.round_to)})"


def operand_expression(kind: str, value: str) -> str:
    """Render a builder operand as formula text."""
    if kind == "Column":
        return f"[{value}]"
    return str(value).strip()


def build_binary_formula(operator: str, left: str, right: str) -> str:
    template = BUILDER_OPERATORS.get(operator)
    if template is None:
        raise FormulaError(f"Unknown operator: {operator}")
    return template.format(left=f"({left})", right=f"({right})")


def build_multi_formula(operator: str, columns: Sequence[str]) -> str:
    template = MULTI_OPERATORS.get(operator)
    if template is None:
        raise FormulaError(f"Unknown operator: {operator}")
    return template.format(args=", ".join(f"[{c}]" for c in columns))


@dataclass
class OperationResult:
    frame: pd.DataFrame
    applied: List[str] = field(default_factory=list)
    errors: List[Tuple[str, str]] = field(default_factory=list)
    previews: Dict[str, pd.Series] = field(default_factory=dict)


def apply_operations(
    frame: pd.DataFrame, operations: Sequence[Operation]
) -> OperationResult:
    """Apply operations in order.

    Each new column is available to later operations, so calculations can be
    chained the way they would be in adjacent spreadsheet columns.
    """
    out = frame.copy()
    result = OperationResult(frame=out)

    for operation in operations:
        name = operation.output_name.strip()
        if not name:
            result.errors.append(("(unnamed)", "Give the new column a name"))
            continue
        try:
            values = evaluate_formula(out, operation.compiled())
        except FormulaError as exc:
            result.errors.append((name, str(exc)))
            continue
        except Exception as exc:  # noqa: BLE001 - surfaced to the user verbatim
            result.errors.append((name, f"{type(exc).__name__}: {exc}"))
            continue

        out[name] = values
        result.applied.append(name)
        result.previews[name] = values.head(5)

    result.frame = out
    return result
