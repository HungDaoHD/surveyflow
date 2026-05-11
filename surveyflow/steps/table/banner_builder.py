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
    group_label:     str        # e.g. "Q1 x Q2 x Q3" — sig test grouping key
    subgroup_label:  str        # leaf/detail label, always shown in DETAIL row
    letter:          str        # A, B, C … resets when outermost level changes
    mask:            pd.Series
    is_total:        bool       = False   # True → excluded from sig test
    level_labels:    list[str]  = field(default_factory=list)
    matrix_row_code:  str | None       = None  # single row → paired mode
    matrix_row_codes: list[str] | None = None  # multiple rows → stacked paired mode
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
        """Look up the label for a choice code from metadata."""
        if not q_pos_to_meta:
            return str(code)
        meta = q_pos_to_meta.get(q_ref) or q_pos_to_meta.get(_resolve(q_ref))
        if not meta:
            return str(code)
        langs = meta.get("choices_i18n", {}).get(str(code), {})
        return langs.get("en") or langs.get("vi") or next(iter(langs.values()), str(code))

    def _expand_with_levels(entry: dict, group_label: str) -> list[BannerColumn]:
        """Handle banner entries with ``levels`` (multi-level nesting) or ``show_total``.

        Mirrors the JS ``buildBannerTree`` / ``buildLevelChildren`` / ``getLeafs``
        logic from *datatable-editor.html* so that Python and the browser preview
        produce identical column structures.

        datatable.json format
        ---------------------
        ::

            {
                "label":      "Store Type",
                "question":   "Q2",
                "groups":     [{"label": "Super", "value": 1}, ...],   # Lv1
                "show_total": true,                                      # add Total before Lv1
                "levels": [
                    {
                        "question":   "Q3",
                        "label":      "City",
                        "groups":     [{"label": "HCM", "value": 1}, ...],  # Lv2
                        "show_total": true                               # add Total before Lv2
                    },
                    ...   # Lv3, Lv4, Lv5 …
                ]
            }

        Tree → ``BannerColumn`` mapping
        --------------------------------
        * depth 0 (root item label)         → ``group_label``
        * depth 1 … N-1 (intermediate)      → ``level_labels[0 … N-2]``
        * depth N (leaf / deepest group)    → ``subgroup_label``

        The mask is the intersection of all ancestor group filters.
        """
        lv1_q       = entry.get("question", "")
        lv1_col     = _resolve(lv1_q) if lv1_q else ""
        lv1_groups  = entry.get("groups", [])
        levels      = entry.get("levels", [])
        show_total1 = entry.get("show_total", False)

        result:  list[BannerColumn] = []
        counter: list[int]          = [0]

        def _recurse(
            parent_mask: pd.Series,
            remaining:   list[dict],
            path:        list[str],
        ) -> None:
            """
            parent_mask : cumulative respondent filter
            remaining   : level items still to be processed
            path        : labels from Lv1 to current node (root excluded)
            """
            if not remaining:
                # Leaf — emit one BannerColumn
                result.append(BannerColumn(
                    group_label    = group_label,
                    subgroup_label = path[-1] if path else (entry.get("label") or group_label),
                    letter         = _letter(counter[0]),
                    mask           = parent_mask,
                    is_total       = False,
                    level_labels   = path[:-1] if path else [],
                ))
                counter[0] += 1
                return

            lv     = remaining[0]
            rest   = remaining[1:]
            lq     = lv.get("question", "")
            lq_col = _resolve(lq) if lq else ""
            lgrps  = lv.get("groups", [])
            show_t = lv.get("show_total", False)

            # show_total at this level → add a "Total" sibling FIRST (unfiltered)
            if show_t and lgrps:
                _recurse(parent_mask, rest, path + ["Total"])

            if not lgrps:
                # Pass-through: no filter at this level
                _recurse(parent_mask, rest, path + [lv.get("label") or lq])
            else:
                for g in lgrps:
                    g_mask = parent_mask & _make_mask(
                        lq, lq_col, g.get("value"), g.get("values")
                    )
                    _recurse(g_mask, rest, path + [g["label"]])

        all_rows = pd.Series(True, index=df.index)

        # Lv1 — show_total: add "Total" FIRST using rows where banner question is not null
        if show_total1 and lv1_groups:
            not_null_mask = df[lv1_col].notna()
            _recurse(not_null_mask, levels, ["Total"])

        if not lv1_groups:
            # No Lv1 groups → single pass-through node
            _recurse(all_rows, levels, [])
        else:
            for g in lv1_groups:
                g_mask = _make_mask(lv1_q, lv1_col, g.get("value"), g.get("values"))
                _recurse(g_mask, levels, [g["label"]])

        return result

    # If any question-level entry has show_total=True, each question already produces
    # its own Total sub-column → suppress the global Total to avoid a redundant column.
    # Mirrors the same hasShowTotal logic in datatable-editor.html buildBannerTree.
    _has_show_total = any(e.get("show_total") for e in config.get("banner", []))

    columns: list[BannerColumn] = []

    for entry in config.get("banner", []):
        # Prefix group_label with question label when a single question is referenced
        if "question" in entry:
            group_label = f"{entry['question']} - {entry['label']}"
        else:
            group_label = entry["label"]

        # ── Total ──────────────────────────────────────────────────────────────
        # Detect Total: no "groups", no "question", no "cross" key.
        # Suppressed when any question-level entry uses show_total (each has its own Total).
        if "groups" not in entry and "question" not in entry and "cross" not in entry:
            if not _has_show_total:
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

        # ── Matrix rows as banner columns (via "use_matrix_rows") ────────────────
        # Format:
        #   { "label": "Brand", "question": "Q13_1", "use_matrix_rows": true }
        # Generates one BannerColumn per row in choices_i18n.rows.
        # Mask: respondents where Q13_1_r{code} is not NaN (have any response for that row).
        if entry.get("use_matrix_rows"):
            q_ref = entry["question"]
            q     = _resolve(q_ref)
            meta  = None
            if q_pos_to_meta:
                meta = q_pos_to_meta.get(q_ref) or q_pos_to_meta.get(q)

            rows: dict = {}
            if meta:
                rows = meta.get("choices_i18n", {}).get("rows", {})

            letter_idx = 0
            for row_code, label_raw in rows.items():
                if isinstance(label_raw, dict):
                    row_label = label_raw.get("vi") or label_raw.get("en") or row_code
                else:
                    row_label = str(label_raw)

                col_name = f"{q}_r{row_code}"
                if col_name in df.columns:
                    mask = df[col_name].notna()
                else:
                    mask = pd.Series(False, index=df.index)

                columns.append(BannerColumn(
                    group_label=group_label,
                    subgroup_label=row_label,
                    letter=_letter(letter_idx),
                    mask=mask,
                    is_total=False,
                ))
                letter_idx += 1
            continue

        # ── Multi-level / show_total banner (via "levels" or "show_total" key) ──────
        # Triggered when the entry declares:
        #   "levels": [...]   → multi-level nesting (Lv2, Lv3, …)
        #   "show_total": true → add a Total sub-column before the groups
        # Uses the recursive _expand_with_levels helper (mirrors JS buildBannerTree).
        if "levels" in entry or entry.get("show_total"):
            columns.extend(_expand_with_levels(entry, group_label))
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


def nest_banner_with_matrix_rows(
    base_columns: list[BannerColumn],
    matrix_question: str,
    df: pd.DataFrame,
    q_pos_to_meta: dict | None = None,
    col_map: dict | None = None,
    groups: list[dict] | None = None,
) -> list[BannerColumn]:
    """Expand each base banner column into sub-columns — one per matrix row (or group).

    Parameters
    ----------
    groups
        Optional explicit group list. Each entry is a dict with:
          - ``label``     : display label for the column header
          - ``row_code``  : single row code  (enables paired-matrix stub mode)
          - ``row_codes`` : list of row codes to union (mask = any row not null;
                            paired-matrix stub mode is disabled for these columns)
        When omitted, every row in ``choices_i18n.rows`` becomes its own column.

    Header levels produced
    ----------------------
    - Total column  → group / brand        [level_labels=[]]
    - Other columns → group / sub-group / brand  [level_labels=[sub-group]]

    ``matrix_row_code`` is set only for single-row columns; multi-row (grouped)
    columns get ``matrix_row_code=None`` and fall back to standard stub expansion.
    """

    def _resolve(q: str) -> str:
        return col_map[q] if col_map and q in col_map else q

    q    = _resolve(matrix_question)
    meta = None
    if q_pos_to_meta:
        meta = q_pos_to_meta.get(matrix_question) or q_pos_to_meta.get(q)

    if not meta:
        return base_columns

    rows: dict = meta.get("choices_i18n", {}).get("rows", {})

    # ── Build brand_items as list of dicts ───────────────────────────────────
    # Each dict: { label, mask, row_code (single), row_codes (grouped) }
    brand_items: list[dict] = []

    if groups:
        for grp in groups:
            label = grp["label"]
            if "row_code" in grp:
                rc       = str(grp["row_code"])
                col_name = f"{q}_r{rc}"
                mask     = df[col_name].notna() if col_name in df.columns \
                           else pd.Series(False, index=df.index)
                brand_items.append({"label": label, "mask": mask,
                                    "row_code": rc, "row_codes": None})
            else:
                rcs  = [str(c) for c in grp.get("row_codes", [])]
                mask = pd.Series(False, index=df.index)
                for rc in rcs:
                    col_name = f"{q}_r{rc}"
                    if col_name in df.columns:
                        mask = mask | df[col_name].notna()
                brand_items.append({"label": label, "mask": mask,
                                    "row_code": None, "row_codes": rcs})
    else:
        if not rows:
            return base_columns
        for row_code, label_raw in rows.items():
            label    = label_raw.get("vi") or label_raw.get("en") or row_code \
                       if isinstance(label_raw, dict) else str(label_raw)
            col_name = f"{q}_r{row_code}"
            mask     = df[col_name].notna() if col_name in df.columns \
                       else pd.Series(False, index=df.index)
            brand_items.append({"label": label, "mask": mask,
                                "row_code": row_code, "row_codes": None})

    # ── Nest each base column × brand_items ───────────────────────────────────
    result: list[BannerColumn] = []

    for col in base_columns:
        level_prefix: list[str] = [] if col.is_total \
                                   else list(col.level_labels) + [col.subgroup_label]
        letter_idx = 0
        for bi in brand_items:
            combined_mask = col.mask & bi["mask"]
            result.append(BannerColumn(
                group_label      = col.group_label,
                subgroup_label   = bi["label"],
                letter           = _letter(letter_idx),
                mask             = combined_mask,
                is_total         = False,
                level_labels     = level_prefix,
                matrix_row_code  = bi["row_code"],
                matrix_row_codes = bi["row_codes"],
            ))
            letter_idx += 1

    return result
