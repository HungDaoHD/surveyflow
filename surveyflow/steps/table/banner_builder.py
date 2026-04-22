"""Build banner column definitions from datatable config."""
from __future__ import annotations

import itertools
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
    group_label:    str        # e.g. "Q1 x Q2 x Q3" — sig test grouping key
    subgroup_label: str        # leaf/detail label, always shown in DETAIL row
    letter:         str        # A, B, C … resets when outermost level changes
    mask:           pd.Series
    is_total:       bool       = False   # True → excluded from sig test
    level_labels:   list[str]  = field(default_factory=list)
    # Header display rules:
    # level_labels = []                 → plain: group(5) / detail(6)
    # level_labels = ["A"]              → 2-level: group(5) / A(6) / detail(7)
    # level_labels = ["A","B","C","D"]  → 5-level: group(5)/A(6)/B(7)/C(8)/D(9)/detail(10)
    #
    # Sig-test grouping: columns are compared within the same
    # (group_label, level_labels[0], …, level_labels[-1]) bucket.


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
        Optional mapping from datatable ``question`` references (``"Q1"``)
        to the actual column name in *df* (the question's ``label``).
        When ``None`` the reference is used as-is.
    q_pos_to_meta
        Optional mapping from question reference → metadata entry dict.
        Used to detect MA questions and apply the correct mask logic,
        and to look up choice labels for ``cross`` banners.
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

    def _code_label(q_ref: str, code: int) -> str:
        """Look up the English label for a choice code from metadata."""
        if not q_pos_to_meta:
            return str(code)
        meta = q_pos_to_meta.get(q_ref) or q_pos_to_meta.get(_resolve(q_ref))
        if not meta:
            return str(code)
        return meta.get("choices_i18n", {}).get(str(code), {}).get("en", str(code))

    columns: list[BannerColumn] = []

    for entry in config.get("banner", []):
        # Prefix group_label with question label when a single question is referenced
        if "question" in entry:
            group_label = f"{entry['question']} - {entry['label']}"
        else:
            group_label = entry["label"]

        # ── Total ──────────────────────────────────────────────────────────────
        # Detect Total: no "groups", no "question", no "cross" key.
        if "groups" not in entry and "question" not in entry and "cross" not in entry:
            columns.append(BannerColumn(
                group_label=group_label,
                subgroup_label="Total",
                letter="",
                mask=pd.Series(True, index=df.index),
                is_total=True,
            ))
            continue

        # ── N-way cross-tab (via "cross" key) ──────────────────────────────────
        # Format:
        #   { "label": "Q1 x Q2 x Q3", "cross": [
        #       { "question": "Q1", "values": [2, 3] },
        #       { "question": "Q2", "values": [2, 3, 4] },
        #       ...
        #   ]}
        # Generates the cartesian product of all dimensions.
        # level_labels = labels for dims[0..n-2]; subgroup_label = label for dim[-1].
        if "cross" in entry:
            dims = entry["cross"]
            dim_items: list[list[dict]] = []
            for dim in dims:
                q_ref = dim["question"]
                q     = _resolve(q_ref)
                items = []
                for code in dim["values"]:
                    label = _code_label(q_ref, code)
                    mask  = _make_mask(q_ref, q, value=code, values=None)
                    items.append({"label": label, "mask": mask})
                dim_items.append(items)

            letter_idx  = 0
            prev_outer: str | None = None

            for combo in itertools.product(*dim_items):
                outer = combo[0]["label"]
                if outer != prev_outer:
                    letter_idx = 0
                    prev_outer = outer

                combined_mask = pd.Series(True, index=df.index)
                for item in combo:
                    combined_mask = combined_mask & item["mask"]

                columns.append(BannerColumn(
                    group_label=group_label,
                    subgroup_label=combo[-1]["label"],
                    letter=_letter(letter_idx),
                    mask=combined_mask,
                    is_total=False,
                    level_labels=[item["label"] for item in combo[:-1]],
                ))
                letter_idx += 1
            continue

        # ── Regular banner / manual cross-tab (via "groups" key) ───────────────
        # Supports:
        #   - Single-question banner:  entry has "question" + groups with value/values
        #   - Manual cross-tab:        groups have "conditions" + optional "subgroup"/"subgroup2"
        letter_idx  = 0
        prev_outer: str | None = None

        for grp in entry.get("groups", []):
            sub_mid = grp.get("subgroup2", "")   # outermost intermediate level
            mid_lbl = grp.get("subgroup",  "")   # inner intermediate level

            level_labels = [x for x in [sub_mid, mid_lbl] if x]
            outer = level_labels[0] if level_labels else None
            if outer and outer != prev_outer:
                letter_idx = 0
                prev_outer = outer

            # ── Build mask ──────────────────────────────────────────────────────
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
                level_labels=level_labels,
            ))
            letter_idx += 1

    return columns
