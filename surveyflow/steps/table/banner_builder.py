"""Build banner column definitions from datatable config."""
from __future__ import annotations

from dataclasses import dataclass, field
import pandas as pd


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
) -> list[BannerColumn]:
    """Return one BannerColumn per banner subgroup defined in config.

    Parameters
    ----------
    col_map
        Optional mapping from datatable ``question`` references (``"q10"``)
        to the actual column name in *df* (the question's ``label``).
        When ``None`` the reference is used as-is.
    """

    def _resolve(q: str) -> str:
        return col_map[q] if col_map and q in col_map else q

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
                    cq = _resolve(cond["question"])
                    if "value" in cond:
                        mask = mask & (df[cq] == cond["value"])
                    elif "values" in cond:
                        mask = mask & df[cq].isin(cond["values"])
            else:
                q = _resolve(entry["question"])
                if "value" in grp:
                    mask = df[q] == grp["value"]
                elif "values" in grp:
                    mask = df[q].isin(grp["values"])
                else:
                    mask = pd.Series(False, index=df.index)

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
