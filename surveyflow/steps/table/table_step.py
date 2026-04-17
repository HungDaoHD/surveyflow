"""TableStep: datatable.json + rawdata + metadata → datatable.xlsx (multi-sheet)."""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from surveyflow.core.base import Step
from surveyflow.steps.table.banner_builder import BannerColumn, build_banner
from surveyflow.steps.table.table_generator import StubBlock, StubRow, compute_table

logger = logging.getLogger(__name__)

# ── Fills ──────────────────────────────────────────────────────────────────────
_F_GROUP   = PatternFill("solid", fgColor="1F4E79")   # navy  — group row
_F_MID     = PatternFill("solid", fgColor="2E75B6")   # blue  — mid / sub_mid rows
_F_DETAIL  = PatternFill("solid", fgColor="BDD7EE")   # light blue — detail row
_F_STAT    = PatternFill("solid", fgColor="EBF3FB")   # very light blue — stat rows
_F_SIG_HDR = PatternFill("solid", fgColor="E2EFDA")   # light green — sig col header

# ── Fonts ──────────────────────────────────────────────────────────────────────
_WHITE_BOLD = Font(color="FFFFFF", bold=True, size=10)
_BOLD       = Font(bold=True, size=10)
_DARK_BOLD  = Font(color="1F4E79", bold=True, size=10)
_RED        = Font(color="FF0000", size=10)
_NORMAL     = Font(size=10)
_GREY       = Font(color="595959", size=9)

# ── Alignments ─────────────────────────────────────────────────────────────────
_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
_L = Alignment(horizontal="left",   vertical="center", wrap_text=True)
_R = Alignment(horizontal="right",  vertical="center")

# ── Borders (black) ────────────────────────────────────────────────────────────
_S_THIN  = Side(style="thin",   color="000000")
_S_MED   = Side(style="medium", color="000000")
_S_THICK = Side(style="thick",  color="000000")

def _brd(thick_left: bool = False, thick_bottom: bool = False) -> Border:
    """Build a cell border with optional thick left (group sep) and thick bottom (section sep)."""
    return Border(
        left   = _S_THICK if thick_left   else _S_THIN,
        right  = _S_THIN,
        top    = _S_THIN,
        bottom = _S_MED   if thick_bottom else _S_THIN,
    )

_BORDER = _brd()   # plain thin border (backwards-compat alias)

_STAT_TYPES = {"t2b", "b2b", "mean", "std", "se"}

# ── Cell helpers ───────────────────────────────────────────────────────────────

def _set(ws, row, col, val=None, *, font=None, fill=None, align=None,
         border=None, num_fmt=None):
    cell = ws.cell(row=row, column=col)
    if val is not None:
        cell.value = val
    if font:    cell.font          = font
    if fill:    cell.fill          = fill
    if align:   cell.alignment     = align
    if border:  cell.border        = border
    if num_fmt: cell.number_format = num_fmt


def _merge(ws, r1, c1, r2, c2):
    if r1 == r2 and c1 == c2:
        return
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


# ── Banner column layout ───────────────────────────────────────────────────────

def _banner_layout(banner_cols: list[BannerColumn], show_sig: bool) -> list[dict]:
    """
    Return a flat list of column-slot dicts, one per Excel data column.

    Fields per slot
    ---------------
    kind           : "data" | "sig"
    banner         : BannerColumn
    bc_index       : int   — position in banner_cols, used to index StubRow dicts
    first_in_group : bool  — True on the first column of each new banner group
                             → render a thick left border to visually separate groups
    """
    slots: list[dict] = []
    prev_group: str | None = None
    for bc_idx, bc in enumerate(banner_cols):
        first = bc.group_label != prev_group
        prev_group = bc.group_label
        slots.append({
            "kind": "data", "banner": bc,
            "bc_index": bc_idx, "first_in_group": first,
        })
        if show_sig:
            slots.append({
                "kind": "sig", "banner": bc,
                "bc_index": bc_idx, "first_in_group": False,
            })
    return slots


# ── Sheet writer ───────────────────────────────────────────────────────────────

def _write_sheet(
    ws,
    config: dict,
    table_cfg: dict,
    banner_cols: list[BannerColumn],
    blocks: list[StubBlock],
) -> None:
    cell_content = table_cfg.get("cell_content", "percentage")   # "count" | "percentage"
    show_sig     = table_cfg.get("show_sig", False)
    sig_cfg      = config.get("significance_test", {"enabled": False})

    slots   = _banner_layout(banner_cols, show_sig)
    n_slots = len(slots)
    DATA_COL = 4   # col D is first data column

    # ── Column widths ────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 30
    for i, slot in enumerate(slots):
        ltr = get_column_letter(DATA_COL + i)
        ws.column_dimensions[ltr].width = 6 if slot["kind"] == "sig" else 13

    # ── Row 1: title ─────────────────────────────────────────────────
    _merge(ws, 1, 1, 1, DATA_COL + n_slots - 1)
    _set(ws, 1, 1, config.get("title", "Data Table"),
         font=Font(bold=True, size=13, color="1F4E79"), align=_L)
    ws.row_dimensions[1].height = 22

    # ── Row 2: spacer ────────────────────────────────────────────────
    ws.row_dimensions[2].height = 5

    # ── Row 3: cell content info ─────────────────────────────────────
    _set(ws, 3, 1, f"Cell content: {cell_content}", font=_NORMAL)
    ws.row_dimensions[3].height = 16

    # ── Row 4: sig test info (method + levels) ───────────────────────
    if show_sig and sig_cfg.get("enabled"):
        lvls   = sig_cfg.get("levels", [95])
        method = sig_cfg.get("method", "independent").capitalize()
        parts  = []
        if 95 in lvls: parts.append("Uppercase = 95%")
        if 90 in lvls: parts.append("lowercase = 90%")
        _set(ws, 4, 1,
             f"Sig test: {', '.join(parts)}    |    Method: {method}",
             font=_NORMAL)
    ws.row_dimensions[4].height = 16

    # ── Row 5: banner GROUP labels ───────────────────────────────────
    group_spans: list[tuple[str, int, int, bool]] = []   # label, c1, c2, first
    for i, slot in enumerate(slots):
        col   = DATA_COL + i
        label = slot["banner"].group_label
        if group_spans and group_spans[-1][0] == label:
            group_spans[-1] = (label, group_spans[-1][1], col, group_spans[-1][3])
        else:
            first = len(group_spans) == 0   # no thick for first group's left edge
            group_spans.append((label, col, col, first))

    for idx, (label, c1, c2, _) in enumerate(group_spans):
        thick = idx > 0   # thick left border between groups (not on the very first)
        _merge(ws, 5, c1, 5, c2)
        _set(ws, 5, c1, label,
             font=_WHITE_BOLD, fill=_F_GROUP, align=_C, border=_brd(thick))
    ws.row_dimensions[5].height = 20

    # ── Dynamic row numbering ────────────────────────────────────────────
    # has_sub_mid → 4-row header: group(5)/sub_mid(6)/mid(7)/detail(8)/letter(9)
    # has_mid only → 3-row header: group(5)/mid(6)/detail(7)/letter(8)
    # plain       → 2-row header: group(5)/detail(6)/letter(7)
    has_sub_mid = any(bc.sub_mid_label for bc in banner_cols)
    has_mid     = any(bc.mid_label     for bc in banner_cols)

    if has_sub_mid:
        SUB_MID_ROW = 6
        MID_ROW     = 7
        DETAIL_ROW  = 8
        LETTER_ROW  = 9
    elif has_mid:
        MID_ROW    = 6
        DETAIL_ROW = 7
        LETTER_ROW = 8
    else:
        DETAIL_ROW = 6
        LETTER_ROW = 7

    DATA_START = (LETTER_ROW + 1) if show_sig else (DETAIL_ROW + 1)

    # ── Row 6 (sub_mid row, only when has_sub_mid) ───────────────────────
    # • 3-level cols (sub_mid_label):    show sub_mid_label, merge by (group, sub_mid)
    # • 2-level cols (mid_label only):   show mid_label,     merge by (group, mid)
    # • regular cols (no mid, no sub):   empty fill,         merge across full group
    if has_sub_mid:
        i = 0
        while i < len(slots):
            slot  = slots[i]
            bc    = slot["banner"]
            c1    = DATA_COL + i
            thick = slot["first_in_group"] and i > 0
            if bc.sub_mid_label:
                j = i
                while (j < len(slots)
                       and slots[j]["banner"].group_label   == bc.group_label
                       and slots[j]["banner"].sub_mid_label == bc.sub_mid_label):
                    j += 1
                _merge(ws, SUB_MID_ROW, c1, SUB_MID_ROW, DATA_COL + j - 1)
                _set(ws, SUB_MID_ROW, c1, bc.sub_mid_label,
                     font=_WHITE_BOLD, fill=_F_GROUP, align=_C, border=_brd(thick))
            elif bc.mid_label:
                # 2-level cross-banner: show mid_label in this outermost row
                j = i
                while (j < len(slots)
                       and slots[j]["banner"].group_label == bc.group_label
                       and slots[j]["banner"].mid_label   == bc.mid_label
                       and not slots[j]["banner"].sub_mid_label):
                    j += 1
                _merge(ws, SUB_MID_ROW, c1, SUB_MID_ROW, DATA_COL + j - 1)
                _set(ws, SUB_MID_ROW, c1, bc.mid_label,
                     font=_WHITE_BOLD, fill=_F_MID, align=_C, border=_brd(thick))
            else:
                # Regular column: empty, merge across same-group non-cross cols
                j = i
                while (j < len(slots)
                       and slots[j]["banner"].group_label == bc.group_label
                       and not slots[j]["banner"].sub_mid_label
                       and not slots[j]["banner"].mid_label):
                    j += 1
                _merge(ws, SUB_MID_ROW, c1, SUB_MID_ROW, DATA_COL + j - 1)
                _set(ws, SUB_MID_ROW, c1, fill=_F_MID, border=_brd(thick))
            i = j
        ws.row_dimensions[SUB_MID_ROW].height = 18

    # ── Mid row (row 6 when has_mid-only, row 7 when has_sub_mid) ────────
    # has_sub_mid sheet:
    #   • 3-level cols: show mid_label, merge by (group, sub_mid, mid)
    #   • 2-level cols: empty (label already shown in sub_mid row), merge per mid group
    #   • regular cols: empty, merge across full group
    # has_mid-only sheet:
    #   • cross-banner cols: show mid_label, merge by (group, mid)
    #   • regular cols: empty, merge across full group
    if has_sub_mid or has_mid:
        i = 0
        while i < len(slots):
            slot  = slots[i]
            bc    = slot["banner"]
            c1    = DATA_COL + i
            thick = slot["first_in_group"] and i > 0

            if has_sub_mid:
                if bc.sub_mid_label and bc.mid_label:
                    # 3-level: show mid_label
                    j = i
                    while (j < len(slots)
                           and slots[j]["banner"].group_label   == bc.group_label
                           and slots[j]["banner"].sub_mid_label == bc.sub_mid_label
                           and slots[j]["banner"].mid_label     == bc.mid_label):
                        j += 1
                    _merge(ws, MID_ROW, c1, MID_ROW, DATA_COL + j - 1)
                    _set(ws, MID_ROW, c1, bc.mid_label,
                         font=_WHITE_BOLD, fill=_F_MID, align=_C, border=_brd(thick))
                elif bc.mid_label:
                    # 2-level in 3-level sheet: empty per mid group
                    j = i
                    while (j < len(slots)
                           and slots[j]["banner"].group_label == bc.group_label
                           and slots[j]["banner"].mid_label   == bc.mid_label
                           and not slots[j]["banner"].sub_mid_label):
                        j += 1
                    _merge(ws, MID_ROW, c1, MID_ROW, DATA_COL + j - 1)
                    _set(ws, MID_ROW, c1, fill=_F_MID, border=_brd(thick))
                else:
                    # Regular: empty, merge across full group
                    j = i
                    while (j < len(slots)
                           and slots[j]["banner"].group_label == bc.group_label
                           and not slots[j]["banner"].sub_mid_label
                           and not slots[j]["banner"].mid_label):
                        j += 1
                    _merge(ws, MID_ROW, c1, MID_ROW, DATA_COL + j - 1)
                    _set(ws, MID_ROW, c1, fill=_F_MID, border=_brd(thick))
            else:
                # has_mid only
                if bc.mid_label:
                    j = i
                    while (j < len(slots)
                           and slots[j]["banner"].group_label == bc.group_label
                           and slots[j]["banner"].mid_label   == bc.mid_label):
                        j += 1
                    _merge(ws, MID_ROW, c1, MID_ROW, DATA_COL + j - 1)
                    _set(ws, MID_ROW, c1, bc.mid_label,
                         font=_WHITE_BOLD, fill=_F_MID, align=_C, border=_brd(thick))
                else:
                    # Regular: empty, merge across full group
                    j = i
                    while (j < len(slots)
                           and slots[j]["banner"].group_label == bc.group_label
                           and not slots[j]["banner"].mid_label):
                        j += 1
                    _merge(ws, MID_ROW, c1, MID_ROW, DATA_COL + j - 1)
                    _set(ws, MID_ROW, c1, fill=_F_MID, border=_brd(thick))
            i = j
        ws.row_dimensions[MID_ROW].height = 18

    # ── Detail row: ALL columns show subgroup_label ───────────────────────
    detail_thick_bot = not show_sig
    i = 0
    while i < len(slots):
        slot  = slots[i]
        bc    = slot["banner"]
        c1    = DATA_COL + i
        j     = i
        while j < len(slots) and slots[j]["banner"] is bc:
            j += 1
        thick = slot["first_in_group"] and i > 0
        _merge(ws, DETAIL_ROW, c1, DETAIL_ROW, DATA_COL + j - 1)
        _set(ws, DETAIL_ROW, c1, bc.subgroup_label,
             font=_DARK_BOLD, fill=_F_DETAIL, align=_C,
             border=_brd(thick, thick_bottom=detail_thick_bot))
        i = j
    ws.row_dimensions[DETAIL_ROW].height = 18

    # ── Letter row ────────────────────────────────────────────────────────
    if show_sig:
        for i, slot in enumerate(slots):
            col   = DATA_COL + i
            thick = slot["first_in_group"] and i > 0
            if slot["kind"] == "data":
                _set(ws, LETTER_ROW, col, slot["banner"].letter,
                     font=_BOLD, fill=_F_DETAIL, align=_C,
                     border=_brd(thick, thick_bottom=True))
            else:
                _set(ws, LETTER_ROW, col, "Sig",
                     font=_GREY, fill=_F_STAT, align=_C,
                     border=_brd(False, thick_bottom=True))
        ws.row_dimensions[LETTER_ROW].height = 16
    else:
        ws.row_dimensions[LETTER_ROW].height = 0

    # ── Stub rows ────────────────────────────────────────────────────
    cur = DATA_START
    for block in blocks:
        if not block.rows:
            continue

        # Separate base row from the rest
        base_row  = next((r for r in block.rows if r.row_type == "base"), None)
        data_rows = [r for r in block.rows if r.row_type != "base"]

        # ── Question header row (doubles as Base row) ────────────────
        # thick_bottom separates this header from its answer-option rows
        _set(ws, cur, 1, block.question_code,
             font=_WHITE_BOLD, fill=_F_MID, align=_C,
             border=_brd(thick_bottom=True))
        _set(ws, cur, 2, block.question_label,
             font=_WHITE_BOLD, fill=_F_MID, align=_L,
             border=_brd(thick_bottom=True))
        _set(ws, cur, 3, "Base" if base_row else "",
             font=_WHITE_BOLD, fill=_F_MID, align=_L,
             border=_brd(thick_bottom=True))

        for i, slot in enumerate(slots):
            col   = DATA_COL + i
            thick = slot["first_in_group"] and i > 0

            if slot["kind"] == "sig" or base_row is None:
                _set(ws, cur, col, fill=_F_MID, border=_brd(thick, thick_bottom=True))
                continue

            bc_idx = slot["bc_index"]
            val    = int(base_row.counts.get(bc_idx, 0))
            _set(ws, cur, col, val,
                 font=_WHITE_BOLD, fill=_F_MID,
                 align=_R, border=_brd(thick, thick_bottom=True), num_fmt="0")

        ws.row_dimensions[cur].height = 17
        cur += 1

        # ── Data rows ────────────────────────────────────────────────
        n_data = len(data_rows)
        for row_idx, stub_row in enumerate(data_rows):
            r           = cur
            is_stat     = stub_row.row_type in _STAT_TYPES
            is_last_row = row_idx == n_data - 1   # thick bottom on last row of block

            # cols A & B — empty placeholder cells
            _set(ws, r, 1, border=_brd(thick_bottom=is_last_row))
            _set(ws, r, 2, border=_brd(thick_bottom=is_last_row))

            # col C — answer/stat label
            _set(ws, r, 3, stub_row.label,
                 font=_DARK_BOLD if is_stat else _NORMAL,
                 fill=_F_STAT if is_stat else None,
                 align=_L, border=_brd(thick_bottom=is_last_row))

            # data cols
            for i, slot in enumerate(slots):
                col    = DATA_COL + i
                bc_idx = slot["bc_index"]
                thick  = slot["first_in_group"] and i > 0

                if slot["kind"] == "sig":
                    if stub_row.row_type == "percent" and show_sig:
                        mark = stub_row.sig_marks.get(bc_idx, "")
                        _set(ws, r, col, mark or None,
                             font=_RED if mark else _NORMAL,
                             align=_C, border=_brd(False, thick_bottom=is_last_row))
                    else:
                        _set(ws, r, col, border=_brd(False, thick_bottom=is_last_row))
                    continue

                raw = (
                    stub_row.counts.get(bc_idx)
                    if cell_content == "count"
                    else stub_row.values.get(bc_idx)
                )

                if cell_content == "count":
                    val = int(raw) if raw is not None else 0
                    fmt = "0"
                    fnt = _DARK_BOLD if is_stat else _NORMAL
                elif stub_row.row_type == "percent":
                    val = float(raw) if raw is not None else 0.0
                    fmt = "0%"
                    fnt = _NORMAL
                else:
                    val = float(raw) if raw is not None else 0.0
                    fmt = "0%" if stub_row.row_type in ("t2b", "b2b") else "0.00"
                    fnt = _DARK_BOLD

                _set(ws, r, col, val,
                     font=fnt,
                     fill=_F_STAT if is_stat else None,
                     align=_R,
                     border=_brd(thick, thick_bottom=is_last_row),
                     num_fmt=fmt)

            ws.row_dimensions[r].height = 15
            cur += 1

    # freeze panes
    ws.freeze_panes = ws.cell(row=DATA_START, column=DATA_COL)


# ── Step ───────────────────────────────────────────────────────────────────────

class TableStep(Step):
    """
    Context inputs
    --------------
    df               : pd.DataFrame
    metadata         : dict
    datatable_config : dict
    output_dir       : str

    Context outputs
    ---------------
    datatable_path   : str
    """

    def run(self, context: dict[str, Any]) -> dict[str, Any]:
        df       = context["df"]
        metadata = context["metadata"]
        config   = context["datatable_config"]
        out_dir  = Path(context.get("output_dir", "."))
        out_dir.mkdir(parents=True, exist_ok=True)

        sig_config = config.get("significance_test", {"enabled": False})

        # Build lookup maps: q{pos} → df column name (label) and → metadata entry
        # Also index by question label (e.g. "Q20") for convenience in datatable.json
        col_map: dict[str, str] = {}
        q_pos_to_meta: dict[str, dict] = {}
        for meta in metadata.get("questions", {}).values():
            pos = meta.get("position")
            if pos is None:
                continue
            key   = f"q{pos}"
            label = meta.get("label") or key
            col_map[key]        = label
            q_pos_to_meta[key]  = meta
            # Allow referencing by label (e.g. "Q20") in addition to "q29"
            if label and label != key:
                col_map[label]       = label
                q_pos_to_meta[label] = meta

        logger.info("Building banner …")
        banner_cols = build_banner(config, df, col_map=col_map)
        logger.info("  → %d banner columns", len(banner_cols))

        logger.info("Computing table …")
        blocks = compute_table(
            stub_configs=config["stub"],
            banner_cols=banner_cols,
            df=df,
            metadata=metadata,
            sig_config=sig_config,
            col_map=col_map,
            q_pos_to_meta=q_pos_to_meta,
        )
        logger.info("  → %d stub blocks", len(blocks))

        tables = config.get("tables", [
            {"sheet": "Table", "cell_content": "percentage", "show_sig": False}
        ])

        wb = Workbook()
        wb.remove(wb.active)   # remove default blank sheet

        for tbl in tables:
            sheet_name = tbl.get("sheet", "Sheet")
            logger.info("Writing sheet: %s", sheet_name)
            ws = wb.create_sheet(title=sheet_name)
            _write_sheet(ws, config, tbl, banner_cols, blocks)

        output_path = str(out_dir / "datatable.xlsx")
        wb.save(output_path)
        logger.info("Saved → %s", output_path)

        context["datatable_path"] = output_path
        return context
