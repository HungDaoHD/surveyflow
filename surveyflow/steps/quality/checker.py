"""Vectorized evaluation of show_condition / contradiction rules against rawdata."""
from __future__ import annotations

from typing import Any

import pandas as pd


# ── Public helpers ─────────────────────────────────────────────────────────────

def build_label_to_q(questions: dict[str, dict]) -> dict[str, dict]:
    """Build {question_label: q_meta} from metadata ``questions`` dict.

    Includes both top-level questions and their matrix sub-questions so that
    translated compound ids like ``"Q10_r1"`` can be resolved during evaluation.
    """
    label_map: dict[str, dict] = {}
    for q in questions.values():
        if "label" in q:
            label_map[q["label"]] = q
        # Matrix sub-questions (nested under sub_questions dict)
        for sq in q.get("sub_questions", {}).values():
            if "label" in sq:
                label_map[sq["label"]] = sq
    return label_map


def eval_condition_vec(
    node: dict,
    df: pd.DataFrame,
    label_to_q: dict[str, dict],
) -> pd.Series:
    """Recursively evaluate a translated condition node against the full DataFrame.

    Returns a boolean Series — True means the condition is met for that row.

    Handles:
    - AND / OR compound nodes
    - Leaf rules with ``codes`` (SA/MA select comparison)
    - Leaf rules with ``value`` (numeric comparison)
    """
    if not node:
        return pd.Series(True, index=df.index)

    if "condition" in node:
        logic = node["condition"].upper()
        rules = node.get("rules") or []
        if not rules:
            return pd.Series(True, index=df.index)
        masks = [eval_condition_vec(r, df, label_to_q) for r in rules]
        result = masks[0]
        for m in masks[1:]:
            result = (result & m) if logic == "AND" else (result | m)
        return result

    # ── Leaf ──────────────────────────────────────────────────────────────
    q_label  = node.get("question", "")
    operator = node.get("operator", "")
    q_meta   = label_to_q.get(q_label)

    if q_meta is None:
        return pd.Series(False, index=df.index)

    atype    = q_meta.get("answer_type", "")
    raw_cols: list[str] = q_meta.get("rawdata_columns") or [q_meta.get("label", q_label)]

    if operator in ("is_not_null", "is_null"):
        has_ans = has_answer_vec(q_meta, df)
        return has_ans if operator == "is_not_null" else ~has_ans

    if operator.startswith("count_"):
        result = _eval_count(node.get("value"), operator, atype, raw_cols, df)
    elif "codes" in node:
        result = _eval_select(node["codes"], operator, atype, raw_cols, df)
    else:
        result = _eval_numeric(node.get("value"), operator, raw_cols, df)

    # ignore_no_data: 0 → condition is False when respondent has no answer
    # (prevents "not_in", "not_equal" etc. from triggering on null values)
    if node.get("ignore_no_data", 0) == 0:
        result = result & has_answer_vec(q_meta, df)

    return result


_MATRIX_TYPES = {"Matrix_SA", "Matrix_MA", "Matrix_NUM"}


def has_answer_vec(q_meta: dict, df: pd.DataFrame) -> pd.Series:
    """True for each respondent that has a non-null, non-empty answer for *q_meta*.

    Matrix questions: True if ANY sub-question row has a value (not all required).
    MA questions: True if any binary column == 1.
    SA / NUM: True if the single column has a non-empty value.
    """
    atype    = q_meta.get("answer_type", "")
    raw_cols: list[str] = q_meta.get("rawdata_columns") or [q_meta.get("label", "")]

    # ── Matrix (parent question): ANY sub-question row with data = answered ──
    if atype in _MATRIX_TYPES:
        sub_questions = q_meta.get("sub_questions", {})
        sub_cols: list[str] = []

        for sq in sub_questions.values():
            sq_cols  = sq.get("rawdata_columns") or [sq.get("label", "")]
            sq_atype = sq.get("answer_type", "SA")
            if sq_atype == "MA":
                from surveyflow.steps.table.banner_builder import _ma_col_map
                cm = _ma_col_map(sq_cols)
                sub_cols.extend(c for c in cm.values() if c in df.columns)
            else:
                sub_cols.extend(c for c in sq_cols if c in df.columns)

        if not sub_cols:
            # Fallback: use parent rawdata_columns (may list all sub-cols)
            sub_cols = [c for c in raw_cols if c in df.columns]

        if not sub_cols:
            return pd.Series(False, index=df.index)

        return df[sub_cols].notna().any(axis=1)

    # ── MA (binary columns) ────────────────────────────────────────────────
    if atype == "MA":
        from surveyflow.steps.table.banner_builder import _ma_col_map
        cm         = _ma_col_map(raw_cols)
        valid_cols = [c for c in cm.values() if c in df.columns]
        if not valid_cols:
            return pd.Series(False, index=df.index)
        return df[valid_cols].notna().any(axis=1)

    # ── SA / ranking / NUM / sub-question (single column) ─────────────────
    col = raw_cols[0] if raw_cols else ""
    if not col or col not in df.columns:
        return pd.Series(False, index=df.index)
    s = df[col]
    return s.notna() & (s.astype(str).str.strip().isin(["", "nan"]) == False)  # noqa: E712


# ── Private helpers ────────────────────────────────────────────────────────────

def _eval_select(
    codes: list[int],
    operator: str,
    atype: str,
    raw_cols: list[str],
    df: pd.DataFrame,
) -> pd.Series:
    if not raw_cols:
        return pd.Series(False, index=df.index)

    if atype == "MA":
        from surveyflow.steps.table.banner_builder import _ma_col_map
        cm = _ma_col_map(raw_cols)
        present_masks = [
            pd.to_numeric(df[cm[str(c)]], errors="coerce") == 1
            for c in codes if str(c) in cm and cm[str(c)] in df.columns
        ]
        if not present_masks:
            return pd.Series(False, index=df.index)
        any_present = present_masks[0]
        for m in present_masks[1:]:
            any_present = any_present | m
        if operator in ("not_in", "not_equal"):
            return ~any_present
        return any_present

    # SA / ranking / etc.
    col = raw_cols[0]
    if col not in df.columns:
        return pd.Series(False, index=df.index)
    col_num = pd.to_numeric(df[col], errors="coerce")
    if operator == "equal":
        return col_num == codes[0]
    if operator == "not_equal":
        return col_num != codes[0]
    if operator == "in":
        return col_num.isin(codes)
    if operator == "not_in":
        return ~col_num.isin(codes)
    return pd.Series(False, index=df.index)


def _eval_numeric(
    value: Any,
    operator: str,
    raw_cols: list[str],
    df: pd.DataFrame,
) -> pd.Series:
    if not raw_cols or value is None:
        return pd.Series(False, index=df.index)
    col = raw_cols[0]
    if col not in df.columns:
        return pd.Series(False, index=df.index)
    col_num = pd.to_numeric(df[col], errors="coerce")
    v = float(value)
    mapping = {
        "equal":            col_num == v,
        "not_equal":        col_num != v,
        "greater":          col_num > v,
        "greater_or_equal": col_num >= v,
        "less":             col_num < v,
        "less_or_equal":    col_num <= v,
    }
    return mapping.get(operator, pd.Series(False, index=df.index))


def _eval_count(
    value: Any,
    operator: str,
    atype: str,
    raw_cols: list[str],
    df: pd.DataFrame,
) -> pd.Series:
    """Evaluate count_* operators: compare number of selected MA answers vs threshold."""
    if value is None:
        return pd.Series(False, index=df.index)
    threshold = float(value)

    if atype == "MA":
        from surveyflow.steps.table.banner_builder import _ma_col_map
        cm         = _ma_col_map(raw_cols)
        valid_cols = [c for c in cm.values() if c in df.columns]
        if not valid_cols:
            return pd.Series(False, index=df.index)
        # Count number of binary cols == 1 per respondent
        count = (df[valid_cols].apply(pd.to_numeric, errors="coerce") == 1).sum(axis=1)
    else:
        # SA / NUM: treat as single-value count (0 or 1 based on has_answer)
        col = raw_cols[0] if raw_cols else ""
        if col not in df.columns:
            return pd.Series(False, index=df.index)
        count = pd.to_numeric(df[col], errors="coerce").notna().astype(int)

    op = operator.replace("count_", "")
    mapping = {
        "equal":            count == threshold,
        "not_equal":        count != threshold,
        "less":             count < threshold,
        "less_or_equal":    count <= threshold,
        "greater":          count > threshold,
        "greater_or_equal": count >= threshold,
    }
    return mapping.get(op, pd.Series(False, index=df.index))


def describe_condition(node: dict) -> str:
    """Return a compact human-readable string for a condition node."""
    if not node:
        return ""
    if "condition" in node:
        op    = node["condition"]
        parts = [describe_condition(r) for r in (node.get("rules") or [])]
        return f" {op} ".join(f"({p})" if " " in p else p for p in parts)
    q  = node.get("question", "?")
    op = node.get("operator", "?")
    if "codes" in node:
        return f"{q} {op} {node['codes']}"
    return f"{q} {op} {node.get('value')}"


def describe_trigger(
    node: "dict | list | None",
    df: pd.DataFrame,
    label_to_q: dict[str, dict],
) -> str:
    """Return a compact string describing only the sub-conditions that are True for *df*.

    Designed for single-row DataFrames.  Walks the condition tree and returns
    only the OR/AND branches that actually evaluated to True, making it easy to
    see *which* specific rule fired for a given respondent.

    - OR node  → include only the True children
    - AND node → include all children (all must be True for AND to be True)
    - Leaf     → delegate to describe_condition
    """
    if not node:
        return ""

    if isinstance(node, list):
        # Top-level list of rules treated as implicit AND
        parts: list[str] = []
        for r in node:
            try:
                if eval_condition_vec(r, df, label_to_q).all():
                    part = describe_trigger(r, df, label_to_q)
                    if part:
                        parts.append(part)
            except Exception:
                pass
        return " AND ".join(f"({p})" if " " in p else p for p in parts)

    if "condition" in node:
        logic = node["condition"].upper()
        rules = node.get("rules") or []
        parts = []
        for r in rules:
            try:
                if eval_condition_vec(r, df, label_to_q).all():
                    part = describe_trigger(r, df, label_to_q)
                    if part:
                        parts.append(part)
            except Exception:
                pass
        if not parts:
            return ""
        sep = f" {logic} "
        return sep.join(f"({p})" if " " in p else p for p in parts)

    # Leaf node — always included by caller (they already verified it's True)
    return describe_condition(node)
