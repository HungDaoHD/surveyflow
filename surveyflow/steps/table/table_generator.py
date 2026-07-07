"""Compute cross-tabulation blocks for the data table."""
from __future__ import annotations

from dataclasses import dataclass, field
import logging
import re
import warnings
import numpy as np
import pandas as pd
from scipy import stats as _scipy_stats

logger = logging.getLogger(__name__)

from surveyflow.steps.table.banner_builder import BannerColumn, _ma_col_map

CODEABLE_TYPES  = {"SA", "MA", "ranking"}
MATRIX_TYPES    = {"Matrix_SA", "Matrix_MA", "Matrix_NUM"}

_SUB_Q_RE = re.compile(r'^(.+?)_r(\d+)$', re.IGNORECASE)

STAT_LABELS: dict[str, str] = {
    "base": "Base",
    "t2b":  "T2B",
    "b2b":  "B2B",
    "nps":  "NPS",
    "mean": "Mean",
    "std":  "Std",
    "se":   "Standard Error",
    "min":  "Min",
    "max":  "Max",
}

DEFAULT_STATS_ORDER = ["base", "percent", "t2b", "b2b", "nps", "mean", "std", "se", "min", "max"]

NUMERIC_TYPES = {"SA", "NUM", "multiplenumber"}


def _is_unified_choices(choices_cfg: list[dict]) -> bool:
    """Return True when the 'choices' array contains at least one group entry
    (identified by the presence of a ``"codes"`` key with a list value).

    When True the renderer iterates the array in order, treating entries with
    ``"codes"`` as group summaries and entries with ``"code"`` as individual
    rows.  When False the legacy ``groups`` + individual-``choices`` paths are used.
    """
    return any("codes" in e and isinstance(e.get("codes"), list) for e in choices_cfg)


def _extract_factors_from_choices(choices_cfg: list[dict]) -> dict[str, float]:
    """Build a factor map {code_str: factor} from per-choice ``"factor"`` fields.

    Returns an empty dict when no choices carry a ``"factor"`` key (caller
    should fall back to the legacy top-level ``"factors"`` / ``"mean_factor"``
    dict).
    """
    return {
        str(e["code"]): float(e["factor"])
        for e in choices_cfg
        if "code" in e and "factor" in e
    }


def _render_unified_choices_rows(
    choices_cfg: list[dict],
    sorted_codes: list[str],
    avail_codes: "set[str]",
    pct_row_fn: "Any",
    grp_row_fn: "Any",
) -> "list[StubRow]":
    """Render percent rows for a unified choices config.

    Rules:
    1. ALL individual codes from ``sorted_codes`` are shown first, in sorted
       order — EXCEPT those marked ``"hidden": true`` via a ``"code"`` entry.
    2. Group entries (``"codes"`` plural) appear after all individual codes,
       in the order they are listed in ``choices_cfg``.
       - ``"type": "netted"`` → group summary row + sub-codes indented below.
       - Any other type (default ``"combine"``) → group summary row only.
    3. ``"code"`` (singular) entries carry metadata only (``factor``, ``hidden``);
       they do NOT change the display position of individual codes.
    """
    hidden_codes: set[str] = {
        str(e["code"]) for e in choices_cfg
        if "code" in e and e.get("hidden", False)
    }

    result: list[StubRow] = []

    # 1. All individual codes in sorted order, skip hidden
    for code in sorted_codes:
        if code not in hidden_codes:
            result.append(pct_row_fn(code))

    # 2. Group entries in choices_cfg order
    for entry in choices_cfg:
        if "codes" not in entry:
            continue
        if entry.get("hidden", False):
            continue
        gc = [str(c) for c in entry.get("codes", []) if str(c) in avail_codes]
        if not gc:
            continue
        result.append(grp_row_fn(gc, entry["label"]))
        if entry.get("type", "combine") == "netted":
            for c in gc:
                result.append(pct_row_fn(c))

    return result


def _resolve_display_codes(all_codes: list[str], sc_choices: list[dict]) -> list[str]:
    """Return codes in display order, with hidden codes excluded.

    Parameters
    ----------
    all_codes
        All available code strings, pre-sorted in default display order.
    sc_choices
        Optional list of ``{"code": int | str, "hidden"?: bool}`` from the stub
        config's ``"choices"`` key.  When empty / absent → return *all_codes*
        unchanged (default sort, all visible).

    Codes listed in *sc_choices* are placed first (in listed order); codes not
    mentioned are appended at the end in their original default order.
    Hidden codes (``"hidden": true``) are excluded from the result entirely.
    """
    if not sc_choices:
        return list(all_codes)
    code_set = set(all_codes)
    ordered: list[str] = []
    seen: set[str] = set()
    for entry in sc_choices:
        code = str(entry.get("code", ""))
        if code not in code_set:
            continue
        seen.add(code)
        if not entry.get("hidden", False):
            ordered.append(code)
    # Append codes not mentioned in sc_choices (at end, in default order)
    ordered += [c for c in all_codes if c not in seen]
    return ordered


# ── Data structures ────────────────────────────────────────────────────────────
# Values and sig_marks are keyed by column INDEX (int) for uniqueness,
# because letters reset to A in every banner group.

@dataclass
class StubRow:
    label: str
    row_type: str
    counts: dict[int, int]          # col_index → raw count
    values: dict[int, float]        # col_index → proportion or stat value
    sig_marks: dict[int, str] = field(default_factory=dict)  # col_index → "BC" / "a"
    code: str | None = None         # numeric code for percent rows; None for stat rows

    def to_dict(self) -> dict:
        # Use ordered arrays (index 0..N-1) instead of dicts to eliminate
        # repeated string keys and reduce JSON file size by ~60-70%.
        n   = max((k + 1 for k in self.values), default=0)
        sig = [self.sig_marks.get(i, "") for i in range(n)]
        out: dict = {
            "label":    self.label,
            "row_type": self.row_type,
            "code":     self.code,
            "values":   [round(float(self.values.get(i, 0.0)), 4) for i in range(n)],
            "counts":   [int(self.counts.get(i, 0)) for i in range(n)],
        }
        if any(sig):
            out["sig_marks"] = sig
        return out


@dataclass
class StubBlock:
    question_code: str
    question_label: str
    answer_type: str
    rows: list[StubRow]

    def to_dict(self) -> dict:
        return {
            "type":        "stub",
            "question":    self.question_code,
            "label":       self.question_label,
            "answer_type": self.answer_type,
            "rows":        [r.to_dict() for r in self.rows],
        }


@dataclass
class RowGroupBlock:
    """A section grouped by matrix rows (e.g. brands), each containing sub-blocks per question."""
    row_label: str              # e.g. "1. Castrol"
    row_code:  str              # e.g. "1"
    sub_blocks: list[StubBlock]

    def to_dict(self) -> dict:
        return {
            "type":       "row_group",
            "row_label":  self.row_label,
            "row_code":   self.row_code,
            "sub_blocks": [b.to_dict() for b in self.sub_blocks],
        }


@dataclass
class RankingBlock:
    """Ranking question — one section per question, sub_blocks per rank position or flat.

    mode = "rank_dist"  → sub_blocks[0] = Rank 1, sub_blocks[1] = Rank 2, …
    mode = "any_rank"   → sub_blocks[0] = flat MA-style (choice ranked at any position)
    """
    question_code:  str
    question_label: str
    answer_type:    str               # always "ranking"
    mode:           str               # "rank_dist" | "any_rank"
    sub_blocks:     list[StubBlock]

    def to_dict(self) -> dict:
        return {
            "type":        "ranking",
            "question":    self.question_code,
            "label":       self.question_label,
            "answer_type": self.answer_type,
            "mode":        self.mode,
            "sub_blocks":  [b.to_dict() for b in self.sub_blocks],
        }


# ── Significance test ──────────────────────────────────────────────────────────

def _binary_array(sub: pd.DataFrame, col: str, code: str, atype: str) -> np.ndarray:
    """Return float 0/1 array: 1 if respondent selected *code*, else 0.

    All callers now pass individual numeric columns (SA, or binary MA/ranking columns)
    so we always use numeric equality. The *atype* parameter is kept for signature
    compatibility but is no longer used for dispatch.
    """
    return (
        pd.to_numeric(sub[col], errors="coerce").fillna(-1) == int(code)
    ).astype(float).to_numpy()


def _compute_sig_marks(
    col: str,
    code: str,
    atype: str,
    df: pd.DataFrame,
    banner_cols: list[BannerColumn],
    sig_config: dict,
) -> dict[int, str]:
    """Paired or independent-samples t-test for proportions.

    sig_config keys
    ---------------
    enabled : bool
    levels  : list[int]  — subset of [90, 95]
    method  : str        — "independent" (default) | "related"

    Rules
    -----
    1. Skip Total columns (is_total=True)
    2. Only compare columns within the same banner group
    3. Letters reset to A within each group
    4. Uppercase = 95 % level, lowercase = 90 % level

    independent
        Welch's two-sample t-test (ttest_ind, equal_var=False).
        Appropriate when the two banner groups have different respondents
        (e.g. Male vs Female).

    related
        Paired t-test (ttest_rel) on the intersection of respondents
        present in *both* banner masks.
        Appropriate when the same respondents can appear in both groups
        (e.g. overlapping demographic segments).
        Groups with no overlapping respondents are skipped.
    """
    result: dict[int, str] = {i: "" for i in range(len(banner_cols))}
    if not sig_config.get("enabled", False):
        return result

    levels  = sig_config.get("levels", [95])
    method  = sig_config.get("method", "independent")   # "independent" | "related"
    alpha95 = 0.05 if 95 in levels else None
    alpha90 = 0.10 if 90 in levels else None

    # ── Performance optimisation ────────────────────────────────────────────
    # Pre-compute a single numeric Series for *this* column (avoid slicing
    # all 1000+ columns of the DataFrame on every banner-pair call).
    code_int   = int(code)
    col_num    = pd.to_numeric(df[col], errors="coerce").fillna(-1)

    # Pre-compute the binary (0/1) array for every non-total banner column.
    bc_arrs: dict[int, np.ndarray] = {}
    for i, bc in enumerate(banner_cols):
        if bc.is_total:
            continue
        bc_arrs[i] = (col_num[bc.mask] == code_int).astype(float).to_numpy()
    # ────────────────────────────────────────────────────────────────────────

    # Group non-total column indices for pairwise comparison.
    # sig_direction=="rows" (default): compare within same demographic group
    # sig_direction=="columns": compare across demographics within the same brand
    sig_direction = sig_config.get("sig_direction", "rows")
    col_letters   = sig_config.get("col_letters")  # {banner_idx: letter} for "columns" mode

    groups: dict[str, list[int]] = {}
    for i, bc in enumerate(banner_cols):
        if bc.is_total:
            continue
        if sig_direction == "columns":
            if getattr(bc, "from_total", False):
                continue
            key = bc.matrix_row_code or (
                "|".join(bc.matrix_row_codes) if bc.matrix_row_codes else bc.subgroup_label
            )
        else:
            key = bc.group_label + ("|" + "|".join(bc.level_labels) if bc.level_labels else "")
        groups.setdefault(key, []).append(i)

    marks: dict[int, list[str]] = {i: [] for i in range(len(banner_cols))}

    # Suppress scipy precision-loss warning once for all pairs (moving
    # warnings.catch_warnings outside the loop avoids repeated context-manager
    # overhead — previously the dominant cost for large banner × code grids).
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message="Precision loss occurred",
            category=RuntimeWarning,
        )
        for grp_indices in groups.values():
            for a in range(len(grp_indices)):
                for b in range(a + 1, len(grp_indices)):
                    i1, i2 = grp_indices[a], grp_indices[b]
                    bc1, bc2 = banner_cols[i1], banner_cols[i2]

                    try:
                        if method == "related":
                            # Paired t-test — only respondents present in BOTH groups.
                            # Need per-respondent arrays aligned on the intersection,
                            # so we slice col_num (single Series) rather than the full df.
                            both_mask = bc1.mask & bc2.mask
                            if int(both_mask.sum()) < 2:
                                continue
                            arr1 = (col_num[both_mask] == code_int).astype(float).to_numpy()
                            arr2 = arr1  # same column → identical values in both contexts
                            t_stat, p_val = _scipy_stats.ttest_rel(arr1, arr2)
                        else:
                            # Independent (Welch's) t-test — reuse pre-computed arrays.
                            arr1 = bc_arrs.get(i1)
                            arr2 = bc_arrs.get(i2)
                            if arr1 is None or arr2 is None:
                                continue
                            if len(arr1) < 2 or len(arr2) < 2:
                                continue
                            t_stat, p_val = _scipy_stats.ttest_ind(
                                arr1, arr2, equal_var=False
                            )
                    except Exception:
                        continue

                    if np.isnan(p_val):
                        continue

                    # winner (higher proportion) gets the loser's letter
                    winner = i1 if t_stat > 0 else i2
                    loser  = i2 if t_stat > 0 else i1
                    loser_letter = (col_letters or {}).get(loser, banner_cols[loser].letter)

                    if alpha95 and p_val < alpha95:
                        marks[winner].append(loser_letter.upper())
                    elif alpha90 and p_val < alpha90:
                        marks[winner].append(loser_letter.lower())

    return {i: "".join(marks[i]) for i in range(len(banner_cols))}


def _compute_sig_for_col(
    col: str,
    codes: list[str],
    df: pd.DataFrame,
    banner_cols: list[BannerColumn],
    sig_config: dict,
) -> dict[str, dict[int, str]]:
    """Vectorized sig test for ALL codes of a single SA column in one pass.

    Faster than calling _compute_sig_marks per code because:
    - col_num computed once (not once per code)
    - col_num[bc.mask] sliced once per banner col (not once per code per banner col)
    - Welch's t-test vectorized across all codes per pair via numpy broadcasting,
      replacing scipy ttest_ind calls (K scipy calls → 1 numpy op per pair)

    Applies to SA questions and Matrix_SA sub-questions where all codes share
    the same rawdata column.  For MA (each code has its own binary column) the
    per-code _compute_sig_marks path is unchanged.

    Returns: {code_str: {banner_idx: mark_str}}
    """
    n_cols  = len(banner_cols)
    _empty  = {c: {i: "" for i in range(n_cols)} for c in codes}
    if not sig_config.get("enabled", False) or not codes:
        return _empty

    levels  = sig_config.get("levels", [95])
    method  = sig_config.get("method", "independent")
    alpha95 = 0.05 if 95 in levels else None
    alpha90 = 0.10 if 90 in levels else None

    # ── Pre-compute once per column ───────────────────────────────────────────
    col_num   = pd.to_numeric(df[col], errors="coerce").fillna(-1)
    code_ints = np.array([int(c) for c in codes], dtype=np.float64)  # (K,)

    # Per banner col: slice col_num (shared across all codes)
    bc_slices: dict[int, np.ndarray] = {
        i: col_num[bc.mask].to_numpy()
        for i, bc in enumerate(banner_cols) if not bc.is_total
    }
    # ─────────────────────────────────────────────────────────────────────────

    sig_direction = sig_config.get("sig_direction", "rows")
    col_letters   = sig_config.get("col_letters")

    groups: dict[str, list[int]] = {}
    for i, bc in enumerate(banner_cols):
        if bc.is_total:
            continue
        if sig_direction == "columns":
            if getattr(bc, "from_total", False):
                continue
            key = bc.matrix_row_code or (
                "|".join(bc.matrix_row_codes) if bc.matrix_row_codes else bc.subgroup_label
            )
        else:
            key = bc.group_label + ("|" + "|".join(bc.level_labels) if bc.level_labels else "")
        groups.setdefault(key, []).append(i)

    K = len(codes)
    marks: dict[str, dict[int, list[str]]] = {c: {i: [] for i in range(n_cols)} for c in codes}

    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Precision loss occurred", category=RuntimeWarning)

        for grp_indices in groups.values():
            for a in range(len(grp_indices)):
                for b in range(a + 1, len(grp_indices)):
                    i1, i2 = grp_indices[a], grp_indices[b]
                    s1 = bc_slices.get(i1)
                    s2 = bc_slices.get(i2)
                    if s1 is None or s2 is None:
                        continue
                    n1, n2 = len(s1), len(s2)
                    if n1 < 2 or n2 < 2:
                        continue

                    if method == "related":
                        # same column → responses identical for both groups →
                        # difference is always 0 → no meaningful test; skip.
                        continue

                    try:
                        # ── Vectorized Welch's t-test across all K codes at once ──
                        # arr_j shape: (n_j, K)  — row=respondent, col=code
                        arr1 = (s1[:, None] == code_ints).astype(np.float64)
                        arr2 = (s2[:, None] == code_ints).astype(np.float64)

                        m1, m2 = arr1.mean(0), arr2.mean(0)       # (K,)
                        v1 = arr1.var(0, ddof=1) if n1 > 1 else np.zeros(K)
                        v2 = arr2.var(0, ddof=1) if n2 > 1 else np.zeros(K)

                        v1n, v2n = v1 / n1, v2 / n2
                        se    = np.sqrt(v1n + v2n)                # (K,)
                        valid = se > 0

                        t_stats = np.where(valid, (m1 - m2) / np.where(valid, se, 1.0), 0.0)

                        # Welch-Satterthwaite degrees of freedom
                        denom = (v1n ** 2 / (n1 - 1)) + (v2n ** 2 / (n2 - 1))
                        df_w  = np.where(
                            denom > 0,
                            (v1n + v2n) ** 2 / np.where(denom > 0, denom, 1.0),
                            1.0,
                        )

                        # scipy.stats.t.sf accepts array x and array df
                        p_vals = np.where(
                            valid,
                            2.0 * _scipy_stats.t.sf(np.abs(t_stats), df=df_w),
                            np.nan,
                        )
                    except Exception:
                        continue

                    for k in range(K):
                        p_val  = p_vals[k]
                        t_stat = t_stats[k]
                        if np.isnan(p_val):
                            continue
                        winner = i1 if t_stat > 0 else i2
                        loser  = i2 if t_stat > 0 else i1
                        letter = (col_letters or {}).get(loser, banner_cols[loser].letter)
                        if alpha95 and p_val < alpha95:
                            marks[codes[k]][winner].append(letter.upper())
                        elif alpha90 and p_val < alpha90:
                            marks[codes[k]][winner].append(letter.lower())

    return {c: {i: "".join(marks[c][i]) for i in range(n_cols)} for c in codes}


def _compute_sig_paired_mode(
    codes: list[str],
    banner_cols: list["BannerColumn"],
    df: pd.DataFrame,
    col_fn,       # callable: BannerColumn -> str | None
    sig_config: dict,
) -> "dict[str, dict[int, str]]":
    """Sig test for paired-mode (matrix_orientation=horizontal) questions.

    Each banner column reads from a *different* data column (col_fn maps bc →
    actual df column).  Only banner columns with a single matrix_row_code are
    tested; grouped row_codes columns are skipped.
    """
    n_cols = len(banner_cols)
    empty  = {c: {i: "" for i in range(n_cols)} for c in codes}
    if not sig_config.get("enabled", False) or not codes:
        return empty

    levels  = sig_config.get("levels", [95])
    method  = sig_config.get("method", "independent")
    alpha95 = 0.05 if 95 in levels else None
    alpha90 = 0.10 if 90 in levels else None
    sig_direction = sig_config.get("sig_direction", "rows")
    col_letters   = sig_config.get("col_letters")

    code_ints = np.array([int(c) for c in codes], dtype=np.float64)
    K = len(codes)

    # Pre-compute per-banner column numeric array (only single row_code cols).
    bc_arrs: dict[int, np.ndarray] = {}
    for i, bc in enumerate(banner_cols):
        if bc.is_total or bc.matrix_row_codes is not None:
            continue
        col = col_fn(bc)
        if col is None or col not in df.columns:
            continue
        bc_arrs[i] = pd.to_numeric(df[col][bc.mask], errors="coerce").fillna(-1).to_numpy()

    # Build comparison groups.
    groups: dict[str, list[int]] = {}
    for i, bc in enumerate(banner_cols):
        if bc.is_total or bc.matrix_row_codes is not None:
            continue
        if i not in bc_arrs:
            continue
        if sig_direction == "columns":
            if getattr(bc, "from_total", False):
                continue
            key = bc.matrix_row_code or bc.subgroup_label
        else:
            key = bc.group_label + ("|" + "|".join(bc.level_labels) if bc.level_labels else "")
        groups.setdefault(key, []).append(i)

    marks: dict[str, dict[int, list[str]]] = {c: {i: [] for i in range(n_cols)} for c in codes}

    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Precision loss occurred", category=RuntimeWarning)
        for grp_indices in groups.values():
            for a in range(len(grp_indices)):
                for b in range(a + 1, len(grp_indices)):
                    i1, i2 = grp_indices[a], grp_indices[b]
                    s1 = bc_arrs.get(i1)
                    s2 = bc_arrs.get(i2)
                    if s1 is None or s2 is None:
                        continue
                    n1, n2 = len(s1), len(s2)
                    if n1 < 2 or n2 < 2:
                        continue

                    try:
                        arr1 = (s1[:, None] == code_ints).astype(np.float64)
                        arr2 = (s2[:, None] == code_ints).astype(np.float64)

                        if method == "related" and n1 == n2:
                            diff   = arr1 - arr2
                            m_diff = diff.mean(0)
                            s_diff = diff.std(0, ddof=1) if n1 > 1 else np.zeros(K)
                            se     = s_diff / np.sqrt(n1)
                            valid  = se > 0
                            t_stats = np.where(valid, m_diff / np.where(valid, se, 1.0), 0.0)
                            df_w    = np.full(K, float(n1 - 1))
                        else:
                            m1, m2 = arr1.mean(0), arr2.mean(0)
                            v1 = arr1.var(0, ddof=1) if n1 > 1 else np.zeros(K)
                            v2 = arr2.var(0, ddof=1) if n2 > 1 else np.zeros(K)
                            v1n, v2n = v1 / n1, v2 / n2
                            se = np.sqrt(v1n + v2n)
                            valid = se > 0
                            t_stats = np.where(valid, (m1 - m2) / np.where(valid, se, 1.0), 0.0)
                            denom = (v1n ** 2 / (n1 - 1)) + (v2n ** 2 / (n2 - 1))
                            df_w  = np.where(denom > 0, (v1n + v2n) ** 2 / np.where(denom > 0, denom, 1.0), 1.0)

                        p_vals = np.where(valid, 2.0 * _scipy_stats.t.sf(np.abs(t_stats), df=df_w), np.nan)
                    except Exception:
                        continue

                    for k in range(K):
                        p_val  = p_vals[k]
                        t_stat = t_stats[k]
                        if np.isnan(p_val):
                            continue
                        winner = i1 if t_stat > 0 else i2
                        loser  = i2 if t_stat > 0 else i1
                        letter = (col_letters or {}).get(loser, banner_cols[loser].letter)
                        if alpha95 and p_val < alpha95:
                            marks[codes[k]][winner].append(letter.upper())
                        elif alpha90 and p_val < alpha90:
                            marks[codes[k]][winner].append(letter.lower())

    return {c: {i: "".join(marks[c][i]) for i in range(n_cols)} for c in codes}


# ── New-format helpers ─────────────────────────────────────────────────────────

def _sub_q_col(sub_meta: dict, fallback: str = "") -> str:
    """Resolve the actual DataFrame column for a matrix sub-question.
    Stored in rawdata_columns[0]; falls back to label if not set."""
    rc = sub_meta.get("rawdata_columns", [])
    return rc[0] if rc else sub_meta.get("label", fallback)


def _build_row_col_map(q_meta: dict) -> dict[str, str]:
    """Map row_index (str) → actual df column, built from sub_questions.rawdata_columns."""
    result: dict[str, str] = {}
    for sm in q_meta.get("sub_questions", {}).values():
        ri = str(sm.get("row_index", ""))
        if ri:
            rc = sm.get("rawdata_columns", [])
            result[ri] = rc[0] if rc else sm.get("label", "")
    return result


def _base_count_ma(sub: pd.DataFrame, rawdata_cols: list) -> int:
    """Base count for binary-format MA: rows where any binary column is non-NaN."""
    cols = [c for c in rawdata_cols if c in sub.columns]
    if not cols:
        return 0
    return int(sub[cols].notna().any(axis=1).sum())


def _code_count_ma(sub: pd.DataFrame, rawdata_cols: list, code: str) -> int:
    """Count rows selecting *code* in binary-format MA (column == 1).
    Uses _ma_col_map for O(1) lookup instead of linear endswith scan."""
    col = _ma_col_map(rawdata_cols).get(code)
    if col is None or col not in sub.columns:
        return 0
    return int((pd.to_numeric(sub[col], errors="coerce") == 1).sum())


# ── Count helpers ──────────────────────────────────────────────────────────────

def _safe_sum(series: pd.Series) -> int:
    """Sum a boolean series to int, safe against empty pd.StringDtype returning ''."""
    if len(series) == 0:
        return 0
    try:
        return int(series.sum())
    except (ValueError, TypeError):
        return 0


def _base_count(sub: pd.DataFrame, col: str) -> int:
    s = sub[col]
    if len(s) == 0:
        return 0
    return int((s.notna() & (s.astype(str).str.strip() != "")).sum())


def _code_count_sc(sub: pd.DataFrame, col: str, code: str) -> int:
    return _safe_sum(pd.to_numeric(sub[col], errors="coerce") == int(code))


def _code_counts_sc_batch(sub: pd.DataFrame, col: str, codes: list[str]) -> dict[str, int]:
    """Count all SA codes in one value_counts pass (vs. one scan per code)."""
    if col not in sub.columns:
        return {c: 0 for c in codes}
    vc = pd.to_numeric(sub[col], errors="coerce").value_counts()
    return {code: int(vc.get(int(code), 0)) for code in codes}


def _code_counts_ma_batch(sub: pd.DataFrame, rawdata_cols: list, codes: list[str]) -> dict[str, int]:
    """Count all MA codes in one vectorized .sum() (vs. one column sum per code)."""
    cm = _ma_col_map(rawdata_cols)
    result = {code: 0 for code in codes}
    pairs = [(code, cm[code]) for code in codes if code in cm and cm[code] in sub.columns]
    if not pairs:
        return result
    sel_cols = [col for _, col in pairs]
    col_sums = (sub[sel_cols].apply(pd.to_numeric, errors="coerce").fillna(0) == 1).sum()
    for code, col in pairs:
        result[code] = int(col_sums[col])
    return result


# ── Ranking-specific count helpers ────────────────────────────────────────────
# Ranking data format: one column per rank position (rawdata_columns[0] = Rank 1,
# rawdata_columns[1] = Rank 2, …). Each cell holds the integer choice code for
# that rank, or NaN when the respondent did not rank to that depth.

def _base_count_rank_at(sub: pd.DataFrame, rank_cols: list[str], n: int) -> int:
    """Count rows where the respondent gave at least *n* rank positions (col n-1 is non-NaN)."""
    if n < 1 or n > len(rank_cols):
        return 0
    col = rank_cols[n - 1]
    if col not in sub.columns:
        return 0
    return int(sub[col].notna().sum())


def _code_count_rank_at(sub: pd.DataFrame, rank_cols: list[str], code: str, position: int) -> int:
    """Count rows where the choice at rank *position* (1-based) equals *code*."""
    if position < 1 or position > len(rank_cols):
        return 0
    col = rank_cols[position - 1]
    if col not in sub.columns:
        return 0
    return int((pd.to_numeric(sub[col], errors="coerce") == int(code)).sum())


def _base_count_rank_any(sub: pd.DataFrame, rank_cols: list[str]) -> int:
    """Count rows where the respondent ranked at least one choice (any rank column non-NaN)."""
    cols = [c for c in rank_cols if c in sub.columns]
    if not cols:
        return 0
    return int(sub[cols].notna().any(axis=1).sum())


def _code_count_rank_any(sub: pd.DataFrame, rank_cols: list[str], code: str) -> int:
    """Count rows where *code* appears at ANY rank position."""
    cols = [c for c in rank_cols if c in sub.columns]
    if not cols:
        return 0
    v = int(code)
    return int(
        (pd.concat(
            [pd.to_numeric(sub[c], errors="coerce") for c in cols],
            axis=1,
        ) == v).any(axis=1).sum()
    )


# ── Main ───────────────────────────────────────────────────────────────────────

def _i18n(d: dict, lang: str, fallback: str = "") -> str:
    """Pick the best label from an i18n dict given a preferred language.

    Priority: lang → "en" → "vi" → first available value → fallback
    """
    if not d:
        return fallback
    val = d.get(lang)
    if val:
        return val
    for alt in ("en", "vi"):
        if alt != lang:
            val = d.get(alt)
            if val:
                return val
    return next((v for v in d.values() if v), fallback)


def _choice_label(choices_i18n: dict, code: str, lang: str = "vi") -> str:
    """Get display label for a choice code in the requested language."""
    entry = choices_i18n.get(str(code), {})
    if not entry:
        return str(code)
    return _i18n(entry, lang, str(code))


def _row_label_text(v: object, lang: str = "vi") -> str:
    if isinstance(v, dict):
        return _i18n(v, lang)
    return str(v)


def _parse_sub_q_ref(q: str, q_pos_to_meta: dict) -> dict | None:
    """Parse 'Q14_r10' syntax → return sub_question meta dict, or None if not a sub-ref.

    Rule 3: when a matrix sub-question is referenced directly (e.g. Q14_r10),
    it behaves as a regular SA/MA question using the parent's columns as choices.
    """
    m = _SUB_Q_RE.match(q)
    if not m:
        return None
    parent_label = m.group(1)
    row_code     = m.group(2)
    parent_meta  = q_pos_to_meta.get(parent_label)
    if parent_meta is None or parent_meta.get("answer_type", "") not in MATRIX_TYPES:
        return None
    for sm in parent_meta.get("sub_questions", {}).values():
        if str(sm.get("row_index", "")) == row_code:
            return sm
    return None


def _validate_row_group(items: list[dict], q_pos_to_meta: dict) -> dict:
    """Validate row_group rules 1 & 2. Returns the shared rows dict on success."""
    if not items:
        raise ValueError("row_group requires at least one item in 'items'.")

    first_rows: dict | None = None
    first_q:    str | None  = None

    for item in items:
        q      = item.get("question", "")
        q_meta = q_pos_to_meta.get(q)

        # Rule 1: all items must be matrix questions
        if q_meta is None:
            raise ValueError(f"row_group: question '{q}' not found in metadata.")
        atype = q_meta.get("answer_type", "")
        if atype not in MATRIX_TYPES:
            raise ValueError(
                f"row_group: question '{q}' has answer_type '{atype}'. "
                f"Only Matrix_SA / Matrix_MA / Matrix_NUM are allowed in row_group. "
                f"Non-matrix questions must be placed outside the row_group. "
                f"To reference a single matrix row alongside non-matrix questions, "
                f"use the '{q}_r{{row_code}}' syntax instead."
            )

        # Rule 2: all items must have identical choices_i18n.rows
        ci   = q_meta.get("choices_i18n", {})
        rows = ci.get("rows", {}) if isinstance(ci, dict) else {}
        if not rows:
            raise ValueError(
                f"row_group: question '{q}' has no choices_i18n.rows defined."
            )
        if first_rows is None:
            first_rows, first_q = rows, q
        elif rows != first_rows:
            raise ValueError(
                f"row_group: question '{q}' has different choices_i18n.rows "
                f"than '{first_q}'. All items must share identical row definitions."
            )

    return first_rows  # type: ignore[return-value]


def _compute_row_group(
    group_cfg: dict,
    banner_cols: list[BannerColumn],
    df: pd.DataFrame,
    sig_config: dict,
    q_pos_to_meta: dict,
    lang: str = "vi",
) -> list[RowGroupBlock]:
    """Compute blocks for a row_group stub entry.

    Returns one RowGroupBlock per shared row (e.g. one per brand),
    each containing one StubBlock per question item.
    """
    items      = group_cfg.get("items", [])
    shared_rows = _validate_row_group(items, q_pos_to_meta)

    # Pre-slice once per banner column for this group.
    sub_dfs: list[pd.DataFrame] = [df[bc.mask] for bc in banner_cols]
    nb = len(banner_cols)
    result: list[RowGroupBlock] = []

    for row_code, row_label_raw in shared_rows.items():
        row_label  = _row_label_text(row_label_raw, lang)
        sub_blocks: list[StubBlock] = []

        for item in items:
            q       = item["question"]
            q_meta  = q_pos_to_meta[q]
            atype   = q_meta["answer_type"]
            sub_atype = "MA" if atype == "Matrix_MA" else "SA"

            ci          = q_meta.get("choices_i18n", {})
            col_choices = ci.get("columns", {}) if isinstance(ci, dict) else {}

            # Find the sub_question for this row
            sub_meta = next(
                (sm for sm in q_meta.get("sub_questions", {}).values()
                 if str(sm.get("row_index", "")) == str(row_code)),
                None,
            )
            if sub_meta is None:
                continue

            sub_col = _sub_q_col(sub_meta, "")
            sub_raw_cols = sub_meta.get("rawdata_columns", [])
            is_sub_ma = sub_atype == "MA"
            if not is_sub_ma and (not sub_col or sub_col not in df.columns):
                continue

            item_label = (
                item.get("label")
                or _i18n(q_meta.get("question_i18n", {}), lang)
                or q
            )
            req_stats = item.get("stats", ["base", "percent"])
            rg_groups = [g for g in item.get("groups", []) if not g.get("hidden", False)]

            sub_bases: dict[int, int] = {
                i: (_base_count_ma(sub_dfs[i], sub_raw_cols) if is_sub_ma
                    else _base_count(sub_dfs[i], sub_col))
                for i in range(nb)
            }

            sub_rows: list[StubRow] = []
            for stat in DEFAULT_STATS_ORDER:
                if stat not in req_stats:
                    continue
                if stat == "base":
                    sub_rows.append(StubRow(
                        label="Base", row_type="base",
                        counts=dict(sub_bases),
                        values={i: float(v) for i, v in sub_bases.items()},
                    ))
                elif stat == "percent" and col_choices:
                    _choices_cfg_rg = item.get("choices", [])
                    _scodes = sorted(col_choices.keys(), key=lambda x: int(x))
                    if is_sub_ma:
                        _batch = [_code_counts_ma_batch(sub_dfs[i], sub_raw_cols, _scodes) for i in range(nb)]
                    else:
                        _batch = [_code_counts_sc_batch(sub_dfs[i], sub_col, _scodes) for i in range(nb)]
                    _sa_sigs = (
                        {} if is_sub_ma
                        else _compute_sig_for_col(sub_col, _scodes, df, banner_cols, sig_config)
                    )

                    def _rg_pct_row(code: str) -> StubRow:
                        cnts_: dict[int, int]   = {i: _batch[i].get(code, 0) for i in range(nb)}
                        pcts_: dict[int, float] = {i: cnts_[i] / sub_bases[i] if sub_bases[i] else 0.0 for i in range(nb)}
                        if is_sub_ma:
                            _bin = _ma_col_map(sub_raw_cols).get(code)
                            sig_ = _compute_sig_marks(_bin, "1", "SA", df, banner_cols, sig_config) if _bin else {}
                        else:
                            sig_ = _sa_sigs.get(code, {})
                        return StubRow(label=_choice_label(col_choices, code, lang), row_type="percent", counts=cnts_, values=pcts_, sig_marks=sig_, code=code)

                    def _rg_grp_row(grp_codes: list[str], label: str) -> StubRow:
                        g_cnts: dict[int, int]   = {i: sum(_batch[i].get(c, 0) for c in grp_codes) for i in range(nb)}
                        g_pcts: dict[int, float] = {i: g_cnts[i] / sub_bases[i] if sub_bases[i] else 0.0 for i in range(nb)}
                        return StubRow(label=label, row_type="group", counts=g_cnts, values=g_pcts)

                    if _is_unified_choices(_choices_cfg_rg):
                        sub_rows.extend(_render_unified_choices_rows(
                            _choices_cfg_rg, _scodes, set(col_choices),
                            _rg_pct_row, _rg_grp_row,
                        ))
                    else:
                        # Legacy: rg_groups first, then display_codes
                        _display_codes = _resolve_display_codes(_scodes, _choices_cfg_rg)
                        for grp in rg_groups:
                            grp_codes = [str(c) for c in grp.get("codes", []) if str(c) in col_choices]
                            if not grp_codes: continue
                            sub_rows.append(_rg_grp_row(grp_codes, grp["label"]))
                            if grp.get("type", "combine") == "netted":
                                for code in grp_codes: sub_rows.append(_rg_pct_row(code))
                        for code in _display_codes: sub_rows.append(_rg_pct_row(code))

            if sub_rows:
                sub_blocks.append(StubBlock(
                    question_code=q.upper(),
                    question_label=item_label,
                    answer_type=atype,
                    rows=sub_rows,
                ))

        if sub_blocks:
            result.append(RowGroupBlock(
                row_label=row_label,
                row_code=row_code,
                sub_blocks=sub_blocks,
            ))

    return result


def compute_table(
    stub_configs: list[dict],
    banner_cols: list[BannerColumn],
    df: pd.DataFrame,
    metadata: dict,
    sig_config: dict,
    col_map: dict[str, str] | None = None,
    q_pos_to_meta: dict[str, dict] | None = None,
    tag: str = "",
    custom_defs: dict[str, list] | None = None,
    lang: str = "vi",
) -> list[StubBlock | RowGroupBlock | RankingBlock]:
    """Compute cross-tabulation blocks.

    Parameters
    ----------
    lang
        Display language for question / choice labels (default ``"vi"``).
        Passed through to :func:`_choice_label` and :func:`_row_label_text`.
    col_map
        Maps ``q{pos}`` reference → actual df column name (label).
    q_pos_to_meta
        Maps ``q{pos}`` reference → metadata entry dict.
    custom_defs
        Reusable filter-group definitions from ``_custom_defs`` array items.
        Keys are group labels; values are lists of choice dicts
        ``{"code": int, "label": str, "filter": {...}}``.
    """

    def _resolve(q: str) -> str:
        return col_map[q] if col_map and q in col_map else q

    def _get_meta(q: str) -> dict | None:
        return q_pos_to_meta.get(q) if q_pos_to_meta else None

    def _clabel(choices: dict, code: str) -> str:
        """Language-aware choice label."""
        return _choice_label(choices, code, lang)

    def _rlabel(v: object) -> str:
        """Language-aware row label."""
        return _row_label_text(v, lang)

    def _qlabel_from_meta(q_meta: dict, fallback: str = "") -> str:
        """Language-aware question label from question_i18n."""
        qi = q_meta.get("question_i18n", {})
        return _i18n(qi, lang) or q_meta.get("label", fallback) or fallback

    def _eval_filter_stub(flt: dict) -> pd.Series:
        """Recursively evaluate a _custom_defs filter against the full df."""
        if not flt:
            return pd.Series(True, index=df.index)
        if "and" in flt:
            mask = pd.Series(True, index=df.index)
            for sub in flt["and"]:
                mask = mask & _eval_filter_stub(sub)
            return mask
        if "or" in flt:
            mask = pd.Series(False, index=df.index)
            for sub in flt["or"]:
                mask = mask | _eval_filter_stub(sub)
            return mask
        # Leaf
        q_ref  = flt.get("question", "")
        codes  = [int(c) for c in flt.get("codes", [])]
        negate = flt.get("negate", False)
        op     = flt.get("op", "any")
        col    = _resolve(q_ref)
        q_meta_f = _get_meta(q_ref)
        atype_f  = q_meta_f.get("answer_type", "") if q_meta_f else ""

        if col in df.columns:
            mask = pd.to_numeric(df[col], errors="coerce").isin(codes)
        elif atype_f == "MA":
            raw_cols_f = q_meta_f.get("rawdata_columns", []) if q_meta_f else []
            cm_f = _ma_col_map(raw_cols_f)
            sub_masks = [
                pd.to_numeric(df[cm_f[str(c)]], errors="coerce") == 1
                for c in codes if str(c) in cm_f and cm_f[str(c)] in df.columns
            ]
            if sub_masks:
                mask = sub_masks[0]
                for m in sub_masks[1:]:
                    mask = (mask & m) if op == "all" else (mask | m)
            else:
                mask = pd.Series(False, index=df.index)
        else:
            mask = pd.Series(False, index=df.index)

        return ~mask if negate else mask

    blocks: list[StubBlock | RowGroupBlock | RankingBlock] = []
    n = len(banner_cols)
    # Pre-slice once — reused for every question and every stat below.
    sub_dfs: list[pd.DataFrame] = [df[bc.mask] for bc in banner_cols]

    total = len(stub_configs)
    _pfx  = f"{tag} — " if tag else ""

    for q_idx, sc in enumerate(stub_configs, 1):

        # ── row_group entry ───────────────────────────────────────────────────
        if sc.get("row_group"):
            grp_q = sc.get("items", [{}])[0].get("question", "row_group") if sc.get("items") else "row_group"
            logger.info("%s[%d/%d] %s (row_group)", _pfx, q_idx, total, grp_q)
            row_blocks = _compute_row_group(
                sc, banner_cols, df, sig_config, q_pos_to_meta or {}, lang=lang
            )
            blocks.extend(row_blocks)
            continue

        # ── custom_ref stub entry ─────────────────────────────────────────────
        if sc.get("type") == "custom_ref":
            ref_name  = sc.get("ref", "")
            choices   = (custom_defs or {}).get(ref_name, [])
            req_stats = sc.get("stats", ["base", "percent"])
            logger.info("%s[%d/%d] custom_ref:%s (%d choices)", _pfx, q_idx, total, ref_name, len(choices))
            if not choices:
                logger.warning("%s custom_ref '%s' not found in _custom_defs — skipping", _pfx, ref_name)
                continue
            # Base = all respondents per banner column
            cr_bases: dict[int, int] = {i: int(bc.mask.sum()) for i, bc in enumerate(banner_cols)}
            cr_rows: list[StubRow] = []
            for stat in DEFAULT_STATS_ORDER:
                if stat not in req_stats:
                    continue
                if stat == "base":
                    cr_rows.append(StubRow(
                        label="Base", row_type="base",
                        counts=dict(cr_bases),
                        values={i: float(v) for i, v in cr_bases.items()},
                    ))
                elif stat == "percent":
                    for choice in choices:
                        flt    = choice.get("filter", {})
                        c_mask = _eval_filter_stub(flt) if flt else pd.Series(True, index=df.index)
                        cnts: dict[int, int]   = {}
                        pcts: dict[int, float] = {}
                        for i, bc in enumerate(banner_cols):
                            cnt     = int((bc.mask & c_mask).sum())
                            base    = cr_bases[i]
                            cnts[i] = cnt
                            pcts[i] = cnt / base if base else 0.0
                        cr_rows.append(StubRow(
                            label    = choice.get("label", str(choice.get("code", ""))),
                            row_type = "percent",
                            counts   = cnts,
                            values   = pcts,
                            code     = str(choice.get("code", "")),
                        ))
            if cr_rows:
                blocks.append(StubBlock(
                    question_code  = "_custom__",
                    question_label = ref_name,
                    answer_type    = "Custom",
                    rows           = cr_rows,
                ))
            continue

        q = sc.get("question")
        if not q:
            logger.warning("%s[%d/%d] skipping stub entry with no 'question' key: %s", _pfx, q_idx, total, sc)
            continue
        logger.info("%s[%d/%d] %s", _pfx, q_idx, total, q)

        # ── sub-question reference: Q14_r10 syntax ────────────────────────────
        sub_ref = _parse_sub_q_ref(q, q_pos_to_meta or {})
        if sub_ref is not None:
            # Treat the sub-question as a flat SA/MA question
            sub_col      = _sub_q_col(sub_ref, q)
            sub_atype    = sub_ref.get("answer_type", "SA")
            sub_raw_cols = sub_ref.get("rawdata_columns", [])
            is_sub_ma    = sub_atype == "MA"
            col_choices  = sub_ref.get("choices_i18n", {})
            req_stats    = sc.get("stats", ["base", "percent"])

            # Label: datatable config > parent question_i18n + row_label
            parent_label_q = _SUB_Q_RE.match(q).group(1)  # type: ignore[union-attr]
            parent_meta    = (q_pos_to_meta or {}).get(parent_label_q, {})
            default_label  = (
                (_i18n(parent_meta.get("question_i18n", {}), lang) or parent_label_q)
                + f" — {sub_ref.get('row_label', q)}"
            )
            q_label = sc.get("label") or default_label

            if not is_sub_ma and sub_col not in df.columns:
                continue

            sub_bases: dict[int, int] = {
                i: (_base_count_ma(sub_dfs[i], sub_raw_cols) if is_sub_ma
                    else _base_count(sub_dfs[i], sub_col))
                for i in range(n)
            }
            _sub_ref_groups = [g for g in sc.get("groups", []) if not g.get("hidden", False)]
            sub_rows: list[StubRow] = []
            for stat in DEFAULT_STATS_ORDER:
                if stat not in req_stats:
                    continue
                if stat == "base":
                    sub_rows.append(StubRow(
                        label="Base", row_type="base",
                        counts=dict(sub_bases),
                        values={i: float(v) for i, v in sub_bases.items()},
                    ))
                elif stat == "percent" and col_choices:
                    _choices_cfg_sr = sc.get("choices", [])
                    _scodes = sorted(col_choices.keys(), key=lambda x: int(x))
                    if is_sub_ma:
                        _batch = [_code_counts_ma_batch(sub_dfs[i], sub_raw_cols, _scodes) for i in range(n)]
                    else:
                        _batch = [_code_counts_sc_batch(sub_dfs[i], sub_col, _scodes) for i in range(n)]
                    _sa_sigs = (
                        {} if is_sub_ma
                        else _compute_sig_for_col(sub_col, _scodes, df, banner_cols, sig_config)
                    )

                    def _sr_pct_row(code: str) -> StubRow:
                        cnts_: dict[int, int]   = {i: _batch[i].get(code, 0) for i in range(n)}
                        pcts_: dict[int, float] = {i: cnts_[i] / sub_bases[i] if sub_bases[i] else 0.0 for i in range(n)}
                        if is_sub_ma:
                            _bin = _ma_col_map(sub_raw_cols).get(code)
                            sig_ = _compute_sig_marks(_bin, "1", "SA", df, banner_cols, sig_config) if _bin else {}
                        else:
                            sig_ = _sa_sigs.get(code, {})
                        return StubRow(label=_clabel(col_choices, code), row_type="percent", counts=cnts_, values=pcts_, sig_marks=sig_, code=code)

                    def _sr_grp_row(grp_codes: list[str], label: str) -> StubRow:
                        g_cnts: dict[int, int]   = {i: sum(_batch[i].get(c, 0) for c in grp_codes) for i in range(n)}
                        g_pcts: dict[int, float] = {i: g_cnts[i] / sub_bases[i] if sub_bases[i] else 0.0 for i in range(n)}
                        return StubRow(label=label, row_type="group", counts=g_cnts, values=g_pcts)

                    if _is_unified_choices(_choices_cfg_sr):
                        sub_rows.extend(_render_unified_choices_rows(
                            _choices_cfg_sr, _scodes, set(col_choices),
                            _sr_pct_row, _sr_grp_row,
                        ))
                    else:
                        # Legacy: groups first, then display_codes
                        for grp in _sub_ref_groups:
                            grp_codes = [str(c) for c in grp.get("codes", []) if str(c) in col_choices]
                            if not grp_codes: continue
                            sub_rows.append(_sr_grp_row(grp_codes, grp["label"]))
                            if grp.get("type", "combine") == "netted":
                                for code in grp_codes: sub_rows.append(_sr_pct_row(code))
                        _display_codes = _resolve_display_codes(_scodes, _choices_cfg_sr)
                        for code in _display_codes: sub_rows.append(_sr_pct_row(code))
            if sub_rows:
                blocks.append(StubBlock(
                    question_code=q.upper(),
                    question_label=q_label,
                    answer_type=sub_atype,
                    rows=sub_rows,
                ))
            continue

        # ── Regular question ──────────────────────────────────────────────────
        q_col  = _resolve(q)
        q_meta = _get_meta(q)

        if q_meta is None:
            continue

        atype        = q_meta["answer_type"]
        choices_i18n = q_meta.get("choices_i18n", {})
        req_stats    = sc.get("stats", ["base", "percent"])

        # Question label: datatable config > question_i18n["en"] > label > q
        q_label = (
            sc.get("label")
            or _i18n(q_meta.get("question_i18n", {}), lang)
            or q_meta.get("label")
            or q
        )

        # ── Matrix: paired mode (banner cols carry matrix_row_code or matrix_row_codes) ──
        # Activated when banner was built via nest_banner_with_matrix_rows (banner_matrix).
        # Instead of expanding the matrix into N sub-question blocks (one per brand),
        # produce a SINGLE block whose rows are the choice options.
        #
        # Two column variants:
        #   matrix_row_code  (str)        → single brand row: read {q_col}_r{code}
        #   matrix_row_codes (list[str])  → grouped brands:   stack/sum across all _r{rc} cols
        _is_paired = any(
            bc.matrix_row_code is not None or bc.matrix_row_codes is not None
            for bc in banner_cols
        )
        if atype in MATRIX_TYPES and _is_paired:
            sub_atype   = "MA" if atype == "Matrix_MA" else "SA"
            col_choices = choices_i18n.get("columns", {}) if isinstance(choices_i18n, dict) else {}

            # Build row_index → actual df column (supports new format rawdata_columns)
            row_col_map = _build_row_col_map(q_meta)

            def _paired_col(rc: str) -> str:
                """Resolve actual df column for a single matrix row code."""
                return row_col_map.get(rc) or f"{q_col}_r{rc}"

            def _paired_base(sub_df, rc: str) -> int:
                sm = next((s for s in q_meta.get("sub_questions", {}).values()
                           if str(s.get("row_index", "")) == rc), None)
                if sub_atype == "MA" and sm:
                    return _base_count_ma(sub_df, sm.get("rawdata_columns", []))
                col = _paired_col(rc)
                return _base_count(sub_df, col) if col in sub_df.columns else 0

            def _paired_cnt(sub_df, rc: str, code: str) -> int:
                sm = next((s for s in q_meta.get("sub_questions", {}).values()
                           if str(s.get("row_index", "")) == rc), None)
                if sub_atype == "MA" and sm:
                    return _code_count_ma(sub_df, sm.get("rawdata_columns", []), code)
                col = _paired_col(rc)
                return _code_count_sc(sub_df, col, code) if col in sub_df.columns else 0

            paired_bases: dict[int, int] = {}
            for i, bc in enumerate(banner_cols):
                if bc.matrix_row_code is not None:
                    paired_bases[i] = _paired_base(sub_dfs[i], bc.matrix_row_code)
                elif bc.matrix_row_codes is not None:
                    paired_bases[i] = sum(_paired_base(sub_dfs[i], rc)
                                          for rc in bc.matrix_row_codes)
                else:
                    paired_bases[i] = 0

            p_rows: list[StubRow] = []
            for stat in DEFAULT_STATS_ORDER:
                if stat not in req_stats:
                    continue
                if stat == "base":
                    p_rows.append(StubRow(
                        label="Base", row_type="base",
                        counts=dict(paired_bases),
                        values={i: float(v) for i, v in paired_bases.items()},
                    ))
                elif stat == "percent" and col_choices:
                    _p_scodes = sorted(col_choices.keys(), key=lambda x: int(x))
                    _choices_cfg_p = sc.get("choices", [])

                    # Sig test for paired mode: each bc reads a different column.
                    def _p_col_fn(bc) -> str | None:
                        if bc.matrix_row_code is None:
                            return None
                        return row_col_map.get(bc.matrix_row_code) or f"{q_col}_r{bc.matrix_row_code}"
                    _p_sigs = _compute_sig_paired_mode(_p_scodes, banner_cols, df, _p_col_fn, sig_config)

                    def _p_cnt_for_code(code: str) -> dict[int, int]:
                        _cnts: dict[int, int] = {}
                        for _i, _bc in enumerate(banner_cols):
                            if _bc.matrix_row_code is not None:
                                _cnts[_i] = _paired_cnt(sub_dfs[_i], _bc.matrix_row_code, code)
                            elif _bc.matrix_row_codes is not None:
                                _cnts[_i] = sum(_paired_cnt(sub_dfs[_i], rc, code) for rc in _bc.matrix_row_codes)
                            else:
                                _cnts[_i] = 0
                        return _cnts

                    def _p_pct_row(code: str) -> StubRow:
                        cnts_ = _p_cnt_for_code(code)
                        pcts_ = {i: cnts_[i] / paired_bases[i] if paired_bases[i] else 0.0 for i in range(n)}
                        sig_  = _p_sigs.get(code, {})
                        return StubRow(label=_clabel(col_choices, code), row_type="percent", counts=cnts_, values=pcts_, sig_marks=sig_, code=code)

                    def _p_grp_row(grp_codes: list[str], label: str) -> StubRow:
                        g_cnts: dict[int, int]   = {i: sum(_p_cnt_for_code(c).get(i, 0) for c in grp_codes) for i in range(n)}
                        g_pcts: dict[int, float] = {i: g_cnts[i] / paired_bases[i] if paired_bases[i] else 0.0 for i in range(n)}
                        return StubRow(label=label, row_type="group", counts=g_cnts, values=g_pcts)

                    if _is_unified_choices(_choices_cfg_p):
                        p_rows.extend(_render_unified_choices_rows(
                            _choices_cfg_p, _p_scodes, set(col_choices),
                            _p_pct_row, _p_grp_row,
                        ))
                    else:
                        # Legacy: groups first, then display_codes
                        _p_groups = [g for g in sc.get("groups", []) if not g.get("hidden", False)]
                        _p_display_codes = _resolve_display_codes(_p_scodes, _choices_cfg_p)
                        for grp in _p_groups:
                            grp_codes = [str(c) for c in grp.get("codes", []) if str(c) in col_choices]
                            if not grp_codes: continue
                            p_rows.append(_p_grp_row(grp_codes, grp["label"]))
                            if grp.get("type", "combine") == "netted":
                                for code in grp_codes: p_rows.append(_p_pct_row(code))
                        for code in _p_display_codes: p_rows.append(_p_pct_row(code))
            if p_rows:
                blocks.append(StubBlock(
                    question_code  = q.upper(),
                    question_label = q_label,
                    answer_type    = atype,
                    rows           = p_rows,
                ))
            continue

        # ── Matrix: expand into one block per sub-question (row) ─────────────
        if atype in MATRIX_TYPES:
            sub_atype   = "MA" if atype == "Matrix_MA" else "SA"
            col_choices = choices_i18n.get("columns", {}) if isinstance(choices_i18n, dict) else {}
            mx_groups   = [g for g in sc.get("groups", []) if not g.get("hidden", False)]
            for sub_key, sub_meta in q_meta.get("sub_questions", {}).items():
                sub_col      = _sub_q_col(sub_meta, sub_key)
                sub_raw_cols = sub_meta.get("rawdata_columns", [])
                is_sub_ma    = sub_atype == "MA"
                row_label    = sub_meta.get("row_label", sub_col)
                if not is_sub_ma and sub_col not in df.columns:
                    continue
                sub_bases: dict[int, int] = {
                    i: (_base_count_ma(sub_dfs[i], sub_raw_cols) if is_sub_ma
                        else _base_count(sub_dfs[i], sub_col))
                    for i in range(n)
                }
                sub_rows: list[StubRow] = []
                for stat in DEFAULT_STATS_ORDER:
                    if stat not in req_stats:
                        continue
                    if stat == "base":
                        sub_rows.append(StubRow(
                            label="Base", row_type="base",
                            counts=dict(sub_bases),
                            values={i: float(v) for i, v in sub_bases.items()},
                        ))
                    elif stat == "percent" and col_choices:
                        _choices_cfg_mx = sc.get("choices", [])
                        _scodes = sorted(col_choices.keys(), key=lambda x: int(x))
                        if is_sub_ma:
                            _batch = [_code_counts_ma_batch(sub_dfs[i], sub_raw_cols, _scodes) for i in range(n)]
                        else:
                            _batch = [_code_counts_sc_batch(sub_dfs[i], sub_col, _scodes) for i in range(n)]
                        _sa_sigs = (
                            {} if is_sub_ma
                            else _compute_sig_for_col(sub_col, _scodes, df, banner_cols, sig_config)
                        )

                        def _mx_pct_row(code: str) -> StubRow:
                            cnts_: dict[int, int]   = {i: _batch[i].get(code, 0) for i in range(n)}
                            pcts_: dict[int, float] = {i: cnts_[i] / sub_bases[i] if sub_bases[i] else 0.0 for i in range(n)}
                            if is_sub_ma:
                                _bin = _ma_col_map(sub_raw_cols).get(code)
                                sig_ = _compute_sig_marks(_bin, "1", "SA", df, banner_cols, sig_config) if _bin else {}
                            else:
                                sig_ = _sa_sigs.get(code, {})
                            return StubRow(label=_clabel(col_choices, code), row_type="percent", counts=cnts_, values=pcts_, sig_marks=sig_, code=code)

                        def _mx_grp_row(grp_codes: list[str], label: str) -> StubRow:
                            g_cnts: dict[int, int]   = {i: sum(_batch[i].get(c, 0) for c in grp_codes) for i in range(n)}
                            g_pcts: dict[int, float] = {i: g_cnts[i] / sub_bases[i] if sub_bases[i] else 0.0 for i in range(n)}
                            return StubRow(label=label, row_type="group", counts=g_cnts, values=g_pcts)

                        if _is_unified_choices(_choices_cfg_mx):
                            sub_rows.extend(_render_unified_choices_rows(
                                _choices_cfg_mx, _scodes, set(col_choices),
                                _mx_pct_row, _mx_grp_row,
                            ))
                        else:
                            # Legacy: mx_groups first, then display_codes
                            _display_codes = _resolve_display_codes(_scodes, _choices_cfg_mx)
                            for grp in mx_groups:
                                grp_codes = [str(c) for c in grp.get("codes", []) if str(c) in col_choices]
                                if not grp_codes: continue
                                sub_rows.append(_mx_grp_row(grp_codes, grp["label"]))
                                if grp.get("type", "combine") == "netted":
                                    for code in grp_codes: sub_rows.append(_mx_pct_row(code))
                            for code in _display_codes: sub_rows.append(_mx_pct_row(code))

                    elif stat == "nps":
                        # Matrix_SA sub-question NPS (e.g. per-brand 0-10/1-10
                        # recommend/liking scale) — mirrors the flat-SA "nps"
                        # branch above; was previously missing entirely for
                        # Matrix_SA, so "nps" in stats silently produced no
                        # row for any matrix question (see CLAUDE.md Step 3c).
                        if is_sub_ma or not col_choices:
                            continue
                        _choices_cfg_nps = sc.get("choices", [])
                        _nps_p_codes: list[str] = []
                        _nps_d_codes: list[str] = []
                        for _e in _choices_cfg_nps:
                            _etype = _e.get("type", "")
                            if _etype == "promoters":
                                _nps_p_codes.extend(str(c) for c in _e.get("codes", []))
                            elif _etype == "detractors":
                                _nps_d_codes.extend(str(c) for c in _e.get("codes", []))
                        if not _nps_p_codes and sc.get("nps_promoters"):
                            _nps_p_codes = [str(c) for c in sc["nps_promoters"]]
                        if not _nps_d_codes and sc.get("nps_detractors"):
                            _nps_d_codes = [str(c) for c in sc["nps_detractors"]]
                        if not _nps_p_codes and not _nps_d_codes:
                            continue
                        _nps_batch = [
                            _code_counts_sc_batch(sub_dfs[i], sub_col, _nps_p_codes + _nps_d_codes)
                            for i in range(n)
                        ]
                        nps_vals: dict[int, float] = {}
                        for i in range(n):
                            base = sub_bases[i]
                            if base == 0:
                                nps_vals[i] = 0.0
                                continue
                            p_cnt = sum(_nps_batch[i].get(c, 0) for c in _nps_p_codes)
                            d_cnt = sum(_nps_batch[i].get(c, 0) for c in _nps_d_codes)
                            nps_vals[i] = round((p_cnt / base - d_cnt / base) * 100, 2)
                        sub_rows.append(StubRow(
                            label=STAT_LABELS["nps"], row_type="nps",
                            counts={i: 0 for i in nps_vals}, values=nps_vals,
                        ))

                    elif stat in ("mean", "std", "se", "min", "max"):
                        if is_sub_ma:
                            # Matrix_MA: mean = avg number of columns selected per respondent
                            _mx_ma_cols = [c for c in sub_raw_cols if c in df.columns]
                            if not _mx_ma_cols:
                                continue
                            _mx_stat_vals: dict[int, float] = {}
                            for i in range(n):
                                _sub_ma = sub_dfs[i][[c for c in _mx_ma_cols if c in sub_dfs[i].columns]]
                                _base = sub_bases[i]
                                if _base == 0 or _sub_ma.empty:
                                    _mx_stat_vals[i] = 0.0
                                    continue
                                _sels = _sub_ma.apply(pd.to_numeric, errors="coerce").fillna(0).clip(0, 1).sum(axis=1)
                                _arr = _sels.to_numpy()
                                _m = float(_arr.mean())
                                _s = float(_arr.std(ddof=1)) if len(_arr) > 1 else 0.0
                                if stat == "mean":  _mx_stat_vals[i] = round(_m, 2)
                                elif stat == "std": _mx_stat_vals[i] = round(_s, 2)
                                elif stat == "se":  _mx_stat_vals[i] = round(_s / np.sqrt(len(_arr)), 2)
                                elif stat == "min": _mx_stat_vals[i] = round(float(_sels.min()), 2)
                                elif stat == "max": _mx_stat_vals[i] = round(float(_sels.max()), 2)
                            sub_rows.append(StubRow(
                                label=STAT_LABELS[stat], row_type=stat,
                                counts={i: 0 for i in _mx_stat_vals}, values=_mx_stat_vals,
                            ))
                        else:
                            # Matrix_SA: mean of SA rating values per sub-question column
                            if not sub_col or sub_col not in df.columns:
                                continue
                            _cf_mx = _extract_factors_from_choices(sc.get("choices", []))
                            _mf = _cf_mx or sc.get("factors") or sc.get("mean_factor")
                            if not _mf:
                                sub_rows.append(StubRow(
                                    label=STAT_LABELS[stat], row_type=stat,
                                    counts={i: 0 for i in range(n)},
                                    values={i: 0.0 for i in range(n)},
                                ))
                                continue
                            _mx_stat_vals: dict[int, float] = {}
                            for i in range(n):
                                _raw = pd.to_numeric(sub_dfs[i][sub_col], errors="coerce").dropna()
                                if len(_raw) == 0:
                                    _mx_stat_vals[i] = 0.0
                                    continue
                                _num = _raw.map(lambda v: _mf.get(str(int(v)), _mf.get(int(v)))).dropna() if _mf else _raw
                                if len(_num) == 0:
                                    _mx_stat_vals[i] = 0.0
                                    continue
                                _m = float(_num.mean())
                                _s = float(_num.std(ddof=1)) if len(_num) > 1 else 0.0
                                if stat == "mean":  _mx_stat_vals[i] = round(_m, 2)
                                elif stat == "std": _mx_stat_vals[i] = round(_s, 2)
                                elif stat == "se":  _mx_stat_vals[i] = round(_s / np.sqrt(len(_num)), 2)
                                elif stat == "min": _mx_stat_vals[i] = round(float(_num.min()), 2)
                                elif stat == "max": _mx_stat_vals[i] = round(float(_num.max()), 2)
                            sub_rows.append(StubRow(
                                label=STAT_LABELS[stat], row_type=stat,
                                counts={i: 0 for i in _mx_stat_vals}, values=_mx_stat_vals,
                            ))
                if sub_rows:
                    blocks.append(StubBlock(
                        question_code=f"{q.upper()}_R{sub_meta.get('row_index','')}",
                        question_label=f"{q_label} — {row_label}",
                        answer_type=atype,
                        rows=sub_rows,
                    ))
            continue

        # ── Ranking question ──────────────────────────────────────────────────
        if atype == "ranking":
            # Each rank position is a separate column in rawdata_columns:
            #   rawdata_columns[0] = Rank 1 col, [1] = Rank 2 col, …
            rank_cols = q_meta.get("rawdata_columns", [])
            if not rank_cols or rank_cols[0] not in df.columns:
                continue

            mode     = sc.get("ranking_mode", "rank_dist")
            top_n    = sc.get("ranking_top_n", 0)
            sorted_codes = sorted(choices_i18n.keys(), key=lambda x: int(x))
            # max_pos capped at actual number of rank columns available
            max_pos = min(
                top_n if top_n > 0 else len(rank_cols),
                len(rank_cols),
            )

            if mode == "any_rank":
                # ── MA-style: each choice → % ranked at any position ──────
                any_bases: dict[int, int] = {
                    i: _base_count_rank_any(sub_dfs[i], rank_cols)
                    for i in range(n)
                }
                any_rows: list[StubRow] = []
                for stat in DEFAULT_STATS_ORDER:
                    if stat not in req_stats:
                        continue
                    if stat == "base":
                        any_rows.append(StubRow(
                            label="Base", row_type="base",
                            counts=dict(any_bases),
                            values={i: float(v) for i, v in any_bases.items()},
                        ))
                    elif stat == "percent":
                        for code in sorted_codes:
                            cnts: dict[int, int]   = {}
                            pcts: dict[int, float] = {}
                            for i in range(n):
                                base = any_bases[i]
                                cnt  = _code_count_rank_any(sub_dfs[i], rank_cols, code)
                                cnts[i] = cnt
                                pcts[i] = cnt / base if base else 0.0
                            # sig: treat Rank_1 col as representative for the test
                            sig = _compute_sig_marks(
                                rank_cols[0], code, "SA", df, banner_cols, sig_config
                            )
                            any_rows.append(StubRow(
                                label=_clabel(choices_i18n, code), row_type="percent",
                                counts=cnts, values=pcts, sig_marks=sig, code=code,
                            ))
                flat_block = StubBlock(
                    question_code=q.upper(), question_label=q_label,
                    answer_type="ranking", rows=any_rows,
                )
                blocks.append(RankingBlock(
                    question_code=q.upper(), question_label=q_label,
                    answer_type="ranking", mode="any_rank",
                    sub_blocks=[flat_block],
                ))

            else:
                # ── rank_dist: one StubBlock per rank position ────────────
                rank_sub_blocks: list[StubBlock] = []
                for rank_pos in range(1, max_pos + 1):
                    rank_bases: dict[int, int] = {
                        i: _base_count_rank_at(sub_dfs[i], rank_cols, rank_pos)
                        for i in range(n)
                    }
                    rank_rows: list[StubRow] = []
                    for stat in DEFAULT_STATS_ORDER:
                        if stat not in req_stats:
                            continue
                        if stat == "base":
                            rank_rows.append(StubRow(
                                label="Base", row_type="base",
                                counts=dict(rank_bases),
                                values={i: float(v) for i, v in rank_bases.items()},
                            ))
                        elif stat == "percent":
                            for code in sorted_codes:
                                cnts = {}
                                pcts = {}
                                for i in range(n):
                                    base = rank_bases[i]
                                    cnt  = _code_count_rank_at(sub_dfs[i], rank_cols, code, rank_pos)
                                    cnts[i] = cnt
                                    pcts[i] = cnt / base if base else 0.0
                                sig = _compute_sig_marks(
                                    rank_cols[rank_pos - 1], code, "SA", df, banner_cols, sig_config
                                )
                                rank_rows.append(StubRow(
                                    label=_clabel(choices_i18n, code), row_type="percent",
                                    counts=cnts, values=pcts, sig_marks=sig, code=code,
                                ))
                    if rank_rows:
                        rank_sub_blocks.append(StubBlock(
                            question_code=f"RANK{rank_pos}",
                            question_label=f"Rank {rank_pos}",
                            answer_type="ranking",
                            rows=rank_rows,
                        ))
                if rank_sub_blocks:
                    blocks.append(RankingBlock(
                        question_code=q.upper(), question_label=q_label,
                        answer_type="ranking", mode="rank_dist",
                        sub_blocks=rank_sub_blocks,
                    ))
            continue

        # MA questions are always stored as binary columns (e.g. S12_1, S12_2 …).
        ma_raw_cols = q_meta.get("rawdata_columns", []) if atype == "MA" else []
        is_ma       = atype == "MA"

        if not is_ma and q_col not in df.columns:
            continue

        # base counts keyed by column index
        bases: dict[int, int] = {
            i: (_base_count_ma(sub_dfs[i], ma_raw_cols) if is_ma
                else _base_count(sub_dfs[i], q_col))
            for i in range(n)
        }

        # ── FT (freetext) question — dynamic choices auto-discovered from rawdata ──
        # Unique non-empty text values become percent rows (like SA with no predefined codes).
        # Sig test is not available for text values.
        if atype == "FT":
            ft_series = df[q_col].dropna().astype(str).str.strip()
            ft_unique  = sorted(ft_series[ft_series != ""].unique())
            ft_rows: list[StubRow] = []
            for stat in req_stats:
                if stat == "base":
                    ft_rows.append(StubRow(
                        label="Base", row_type="base",
                        counts=dict(bases),
                        values={i: float(v) for i, v in bases.items()},
                    ))
                elif stat == "percent":
                    for val in ft_unique:
                        cnts_ft: dict[int, int]   = {}
                        pcts_ft: dict[int, float] = {}
                        for i in range(n):
                            sub_s   = sub_dfs[i][q_col].dropna().astype(str).str.strip()
                            cnt     = int((sub_s == val).sum())
                            cnts_ft[i] = cnt
                            pcts_ft[i] = cnt / bases[i] if bases[i] else 0.0
                        ft_rows.append(StubRow(
                            label=val, row_type="percent",
                            counts=cnts_ft, values=pcts_ft,
                        ))
            if ft_rows:
                blocks.append(StubBlock(
                    question_code=q.upper(),
                    question_label=q_label,
                    answer_type=atype,
                    rows=ft_rows,
                ))
            continue

        def _sig_for_code(code: str) -> dict:
            """Sig marks for a code, routing to the correct column."""
            if is_ma:
                bin_col = _ma_col_map(ma_raw_cols).get(code)
                if bin_col:
                    return _compute_sig_marks(bin_col, "1", "SA", df, banner_cols, sig_config)
                return {}
            return _compute_sig_marks(q_col, code, atype, df, banner_cols, sig_config)

        rows: list[StubRow] = []

        for stat in DEFAULT_STATS_ORDER:
            if stat not in req_stats:
                continue

            if stat == "base":
                rows.append(StubRow(
                    label="Base", row_type="base",
                    counts=dict(bases),
                    values={i: float(v) for i, v in bases.items()},
                ))

            elif stat == "percent":
                if atype not in CODEABLE_TYPES or not choices_i18n:
                    continue

                _choices_cfg = sc.get("choices", [])
                sorted_codes  = sorted(choices_i18n.keys(), key=lambda x: int(x))

                # Batch: one value_counts (SA) or one .sum() (MA) per banner col.
                if is_ma:
                    _batch = [_code_counts_ma_batch(sub_dfs[i], ma_raw_cols, sorted_codes)
                              for i in range(n)]
                else:
                    _batch = [_code_counts_sc_batch(sub_dfs[i], q_col, sorted_codes)
                              for i in range(n)]

                # SA: pre-compute sig marks for all codes in one vectorized pass.
                # MA keeps per-code path (each code has its own binary column).
                _sa_sigs: dict[str, dict] = (
                    {} if is_ma
                    else _compute_sig_for_col(q_col, sorted_codes, df, banner_cols, sig_config)
                )

                def _cnt_pct_union(target: list[str]) -> tuple[dict, dict]:
                    """Count/pct for a union of codes (group summary rows)."""
                    _cnts: dict[int, int]   = {}
                    _pcts: dict[int, float] = {}
                    for _i in range(n):
                        _base = bases[_i]
                        _cnt  = sum(_batch[_i][c] for c in target)
                        _cnts[_i] = _cnt
                        _pcts[_i] = _cnt / _base if _base else 0.0
                    return _cnts, _pcts

                def _sa_pct_row(code: str) -> StubRow:
                    """Build a percent StubRow for a single code."""
                    cnts_: dict[int, int]   = {i: _batch[i][code] for i in range(n)}
                    pcts_: dict[int, float] = {i: cnts_[i] / bases[i] if bases[i] else 0.0
                                               for i in range(n)}
                    sig_ = _sa_sigs.get(code, {}) if not is_ma else _sig_for_code(code)
                    return StubRow(
                        label=_clabel(choices_i18n, code), row_type="percent",
                        counts=cnts_, values=pcts_, sig_marks=sig_, code=code,
                    )

                if _is_unified_choices(_choices_cfg):
                    def _sa_grp_row(grp_codes: list[str], label: str) -> StubRow:
                        g_cnts, g_pcts = _cnt_pct_union(grp_codes)
                        return StubRow(label=label, row_type="group", counts=g_cnts, values=g_pcts)
                    rows.extend(_render_unified_choices_rows(
                        _choices_cfg, sorted_codes, set(choices_i18n),
                        _sa_pct_row, _sa_grp_row,
                    ))
                else:
                    # ── Legacy mode: groups first, then display_codes ─────────────
                    # Default group type "combine": summary only; "netted": indented codes.
                    stub_groups   = [g for g in sc.get("groups", []) if not g.get("hidden", False)]
                    display_codes = _resolve_display_codes(sorted_codes, _choices_cfg)

                    for grp in stub_groups:
                        grp_codes = [str(c) for c in grp.get("codes", []) if str(c) in choices_i18n]
                        if not grp_codes:
                            continue
                        g_cnts, g_pcts = _cnt_pct_union(grp_codes)
                        rows.append(StubRow(
                            label=grp["label"], row_type="group",
                            counts=g_cnts, values=g_pcts,
                        ))
                        if grp.get("type", "combine") == "netted":
                            for code in grp_codes:
                                rows.append(_sa_pct_row(code))

                    for code in display_codes:
                        rows.append(_sa_pct_row(code))

            elif stat in ("t2b", "b2b"):
                if atype not in CODEABLE_TYPES or not choices_i18n:
                    continue
                custom_key   = "t2b_codes" if stat == "t2b" else "b2b_codes"
                custom_codes = sc.get(custom_key)
                if custom_codes is not None:
                    target = [str(c) for c in custom_codes]
                else:
                    sorted_codes = sorted(choices_i18n.keys(), key=lambda x: int(x))
                    target = sorted_codes[-2:] if stat == "t2b" else sorted_codes[:2]
                cnts: dict[int, int]   = {}
                pcts: dict[int, float] = {}
                for i in range(n):
                    base = bases[i]
                    if is_ma:
                        cnt = sum(_code_count_ma(sub_dfs[i], ma_raw_cols, c) for c in target)
                    else:
                        cnt = sum(_code_count_sc(sub_dfs[i], q_col, c) for c in target)
                    cnts[i] = cnt
                    pcts[i] = cnt / base if base else 0.0
                rows.append(StubRow(
                    label=STAT_LABELS[stat], row_type=stat,
                    counts=cnts, values=pcts,
                ))

            elif stat == "nps":
                if atype not in CODEABLE_TYPES or not choices_i18n:
                    continue
                _choices_cfg_nps = sc.get("choices", [])
                _nps_p_codes: list[str] = []
                _nps_d_codes: list[str] = []
                for _e in _choices_cfg_nps:
                    _etype = _e.get("type", "")
                    if _etype == "promoters":
                        _nps_p_codes.extend(str(c) for c in _e.get("codes", []))
                    elif _etype == "detractors":
                        _nps_d_codes.extend(str(c) for c in _e.get("codes", []))
                # Fallback: explicit nps_promoters / nps_detractors keys on stub entry
                if not _nps_p_codes and sc.get("nps_promoters"):
                    _nps_p_codes = [str(c) for c in sc["nps_promoters"]]
                if not _nps_d_codes and sc.get("nps_detractors"):
                    _nps_d_codes = [str(c) for c in sc["nps_detractors"]]
                if not _nps_p_codes and not _nps_d_codes:
                    continue
                nps_vals: dict[int, float] = {}
                for i in range(n):
                    base = bases[i]
                    if base == 0:
                        nps_vals[i] = 0.0
                        continue
                    if is_ma:
                        p_cnt = sum(_code_count_ma(sub_dfs[i], ma_raw_cols, c) for c in _nps_p_codes)
                        d_cnt = sum(_code_count_ma(sub_dfs[i], ma_raw_cols, c) for c in _nps_d_codes)
                    else:
                        p_cnt = sum(_code_count_sc(sub_dfs[i], q_col, c) for c in _nps_p_codes)
                        d_cnt = sum(_code_count_sc(sub_dfs[i], q_col, c) for c in _nps_d_codes)
                    nps_vals[i] = round((p_cnt / base - d_cnt / base) * 100, 2)
                rows.append(StubRow(
                    label=STAT_LABELS["nps"], row_type="nps",
                    counts={i: 0 for i in nps_vals}, values=nps_vals,
                ))

            elif stat in ("mean", "std", "se", "min", "max"):
                if atype == "MA":
                    # MA mean = average number of codes selected per respondent.
                    # Base = respondents who saw the question (any binary col non-NaN).
                    _ma_bin = [c for c in ma_raw_cols if c in df.columns]
                    if not _ma_bin:
                        continue
                    stat_vals: dict[int, float] = {}
                    for i in range(n):
                        _sub_ma = sub_dfs[i][[c for c in _ma_bin if c in sub_dfs[i].columns]]
                        _base = bases[i]
                        if _base == 0 or _sub_ma.empty:
                            stat_vals[i] = 0.0
                            continue
                        _sels = _sub_ma.apply(pd.to_numeric, errors="coerce").fillna(0).clip(0, 1).sum(axis=1)
                        _arr = _sels.to_numpy()
                        _m = float(_arr.mean())
                        _s = float(_arr.std(ddof=1)) if len(_arr) > 1 else 0.0
                        if stat == "mean":  stat_vals[i] = round(_m, 2)
                        elif stat == "std": stat_vals[i] = round(_s, 2)
                        elif stat == "se":  stat_vals[i] = round(_s / np.sqrt(len(_arr)), 2)
                        elif stat == "min": stat_vals[i] = round(float(_sels.min()), 2)
                        elif stat == "max": stat_vals[i] = round(float(_sels.max()), 2)
                    rows.append(StubRow(
                        label=STAT_LABELS[stat], row_type=stat,
                        counts={i: 0 for i in stat_vals}, values=stat_vals,
                    ))
                elif atype not in NUMERIC_TYPES:
                    continue
                else:
                    if atype == "SA" and not choices_i18n:
                        continue
                    # factors / mean_factor: maps code → numeric weight (SA only).
                    # Priority: per-choice "factor" fields → "factors" dict → "mean_factor" dict.
                    mean_factor: dict | None = None
                    if atype == "SA":
                        _cf = _extract_factors_from_choices(sc.get("choices", []))
                        mean_factor = _cf or sc.get("factors") or sc.get("mean_factor")
                        if not mean_factor:
                            rows.append(StubRow(
                                label=STAT_LABELS[stat], row_type=stat,
                                counts={i: 0 for i in range(n)},
                                values={i: 0.0 for i in range(n)},
                            ))
                            continue
                    stat_vals: dict[int, float] = {}
                    for i in range(n):
                        raw = pd.to_numeric(sub_dfs[i][q_col], errors="coerce").dropna()
                        if len(raw) == 0:
                            stat_vals[i] = 0.0
                            continue
                        if mean_factor:
                            numeric = raw.map(
                                lambda v: mean_factor.get(str(int(v)), mean_factor.get(int(v)))
                            ).dropna()
                        else:
                            numeric = raw
                        if len(numeric) == 0:
                            stat_vals[i] = 0.0
                            continue
                        m = float(numeric.mean())
                        s = float(numeric.std(ddof=1)) if len(numeric) > 1 else 0.0
                        if stat == "mean":
                            stat_vals[i] = round(m, 2)
                        elif stat == "std":
                            stat_vals[i] = round(s, 2)
                        elif stat == "se":
                            stat_vals[i] = round(s / np.sqrt(len(numeric)), 2)
                        elif stat == "min":
                            stat_vals[i] = round(float(numeric.min()), 2)
                        elif stat == "max":
                            stat_vals[i] = round(float(numeric.max()), 2)
                    rows.append(StubRow(
                        label=STAT_LABELS[stat], row_type=stat,
                        counts={i: 0 for i in stat_vals},
                        values=stat_vals,
                    ))

        blocks.append(StubBlock(
            question_code=q.upper(),
            question_label=q_label,
            answer_type=atype,
            rows=rows,
        ))

    return blocks
