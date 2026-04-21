"""Build banner column definitions from datatable config."""
from __future__ import annotations

from dataclasses import dataclass, field
import pandas as pd


def _ma_contains(series: pd.Series, code: str) -> pd.Series:
    """Return boolean mask: True where *code* appears in semicolon-separated MA column."""
    return series.apply(
        lambda v: code in str(v).split(";")
        if pd.notna(v) and str(v).strip() != "" else False
    )


def _letter(i: int) -> str:
    """0→A, 1→B, …, 25→Z, 26→AA, …"""
    letters = []
    n = i + 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        letters.append(chr(65 + r))
    return "".join(reversed(letters))


@dataclass
class BannerColumn:
    group_label:    str        # e.g. "Gender x Age x Occupation" — sig test grouping key
    subgroup_label: str        # innermost label  e.g. "Working" / "Male" / "<30"
    letter:         str        # A, B, C … resets at outermost mid-level boundary
    mask:           pd.Series
    is_total:       bool = False   # True → excluded from sig test
    mid_label:      str  = ""      # 2nd-level sub-header  e.g. "<30" / "Male"
    sub_mid_label:  str  = ""      # 1st-level sub-header  e.g. "Male" (for 3-level cross)
    #
    # Header display rules
    # ─────────────────────────────────────────────────────────────────
    # 1-level  (no mid, no sub_mid):
    #   row 6 = subgroup_label
    #
    # 2-level  (mid only):
    #   row 6 = mid_label  (merged across same group+mid)
    #   row 7 = subgroup_label
    #
    # 3-level  (sub_mid + mid):
    #   row 6 = sub_mid_label  (merged across same group+sub_mid)
    #   row 7 = mid_label      (merged across same group+sub_mid+mid)
    #   row 8 = subgroup_label
    #
    # For sig-test grouping, columns are compared within the same
    # (group_label, sub_mid_label, mid_label) bucket.


def build_banner(
    config: dict,
    df: pd.DataFrame,
    col_map: dict[str, str] | None = None,
    q_pos_to_meta: dict[str, dict] | None = None,
) -> list[BannerColumn]:
    """Return one BannerColumn per banner subgroup defined in config.

    Parameters
    ----------
    col_map
        Optional mapping from datatable ``question`` references (``"q10"``)
        to the actual column name in *df* (the question's ``label``).
        When ``None`` the reference is used as-is.
    q_pos_to_meta
        Optional mapping from question reference → metadata entry dict.
        Used to detect MA questions and apply the correct mask logic.
    """

    def _resolve(q: str) -> str:
        return col_map[q] if col_map and q in col_map else q

    def _is_ma(q_ref: str) -> bool:
        if not q_pos_to_meta:
            return False
        meta = q_pos_to_meta.get(q_ref) or q_pos_to_meta.get(_resolve(q_ref))
        return (meta or {}).get("answer_type") == "MA"

    def _make_mask(q_ref: str, col: str, value: int | None, values: list | None) -> pd.Series:
        """Build respondent mask for one banner group, handling SA and MA."""
        if _is_ma(q_ref):
            if value is not None:
                return _ma_contains(df[col], str(value))
            elif values:
                codes = [str(v) for v in values]
                return df[col].apply(
                    lambda v: any(c in str(v).split(";") for c in codes)
                    if pd.notna(v) and str(v).strip() != "" else False
                )
            return pd.Series(False, index=df.index)
        else:
            if value is not None:
                return df[col] == value
            elif values:
                return df[col].isin(values)
            return pd.Series(False, index=df.index)

    columns: list[BannerColumn] = []

    for entry in config.get("banner", []):
        group_label = entry["label"]

        # ── Total ──────────────────────────────────────────────────────
        # Detect Total: no "groups" key AND no "question" key.
        # Cross-banners have "groups" but no top-level "question" — NOT Total.
        if "groups" not in entry and "question" not in entry:
            columns.append(BannerColumn(
                group_label=group_label,
                subgroup_label="Total",
                letter="",
                mask=pd.Series(True, index=df.index),
                is_total=True,
            ))
            continue

        # ── Letter index — resets at the outermost mid-level boundary ──
        # • 3-level (sub_mid_label): reset when sub_mid_label changes
        # • 2-level (mid_label only): reset when mid_label changes
        # • Regular (no mid): increments continuously
        letter_idx = 0
        prev_outer = None

        for grp in entry.get("groups", []):
            sub_mid = grp.get("subgroup2", "")   # outermost mid-level
            mid_lbl = grp.get("subgroup",  "")   # inner mid-level (or only mid-level)

            outer = sub_mid if sub_mid else mid_lbl
            if outer and outer != prev_outer:
                letter_idx = 0
                prev_outer = outer

            # ── Build mask ────────────────────────────────────────────
            if "conditions" in grp:
                mask = pd.Series(True, index=df.index)
                for cond in grp["conditions"]:
                    cq_ref = cond["question"]
                    cq     = _resolve(cq_ref)
                    mask   = mask & _make_mask(
                        cq_ref, cq,
                        value=cond.get("value"),
                        values=cond.get("values"),
                    )
            else:
                q_ref = entry["question"]
                q     = _resolve(q_ref)
                mask  = _make_mask(
                    q_ref, q,
                    value=grp.get("value"),
                    values=grp.get("values"),
                )

            columns.append(BannerColumn(
                group_label=group_label,
                subgroup_label=grp["label"],
                letter=_letter(letter_idx),
                mask=mask,
                is_total=False,
                mid_label=mid_lbl,
                sub_mid_label=sub_mid,
            ))
            letter_idx += 1

    return columns
