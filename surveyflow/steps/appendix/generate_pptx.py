#!/usr/bin/env python3
"""Generate PowerPoint slides from chart_data.json — template-clone approach.

Instead of building charts from scratch (which never matches the template's
styling), this clones pre-extracted chart-style templates and only swaps in new
data + text. Every visual property — colors, fonts, hidden value axis,
gridlines, gap width, legend, data-label format — is inherited verbatim from the
templates, so the output looks identical to temp.pptx.

The 4 chart styles live as bundled XML in surveyflow/chart_templates/ (extracted
from temp.pptx by extract_chart_templates.py), so the generator is self-contained
— it does NOT need temp.pptx at runtime. To restyle, edit temp.pptx in PowerPoint
and re-run extract_chart_templates.py.

Chart-style roles:
    bar      bar_clustered, horizontal   → "Total" left chart (MA)
    col      col_clustered,  vertical    → breakdown right charts / vertical Total
    donut    doughnut                    → "Total" left chart (SA≤5)
    stacked  col_percentStacked          → combined breakdown right chart

Usage (CLI after pip install surveyflow):
    surveyflow-pptx <chart_data.json> <output.pptx> [options]

Options:
    --templates DIR    Chart-template dir (default: bundled surveyflow/chart_templates)
    --table N          Export only this table index (default: all)
    --start-page N     Starting page number (default: 1)
"""

from __future__ import annotations

import argparse
import copy
import json
import re
import sys
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import qn
    from lxml import etree
except ImportError:
    sys.exit("Missing dependency: pip install python-pptx")


# ── Color palette (text only — chart colors come from the template) ───────────

C_NAVY   = RGBColor(0x1F, 0x4E, 0x79)  # title, section labels, Q-prefix   (#1F4E79)
C_ORANGE = RGBColor(0xED, 0x7D, 0x31)  # footer bar                        (#ED7D31)
C_QDESC  = RGBColor(0x40, 0x40, 0x40)  # Q-label description                (#404040)
C_QBASE  = RGBColor(0x7F, 0x7F, 0x7F)  # (N=XX) and page number            (#7F7F7F)


# ── Slide geometry — exact EMU values from temp.pptx ─────────────────────────
SW = Emu(12192000)   # 13.333"
SH = Emu(6858000)    # 7.500"

# Title
TITLE_L, TITLE_T = Emu(502920),  Emu(256032)
TITLE_W, TITLE_H = Emu(8686800), Emu(640080)

# Q-label (bottom)
Q_LABEL_L, Q_LABEL_T = Emu(502920),  Emu(5961888)
Q_LABEL_W, Q_LABEL_H = Emu(11521440), Emu(512064)

# Page number
PAGE_L, PAGE_T = Emu(11365992), Emu(6473952)
PAGE_W, PAGE_H = Emu(457200),   Emu(274320)

# Orange footer bar
FOOTER_L, FOOTER_T = Emu(0), Emu(6711696)
FOOTER_W, FOOTER_H = Emu(12188952), Emu(146304)

# Section label height ("Total", "By age")
SECTION_H = Emu(292608)         # 0.32"
SECTION_TO_CHART = Emu(228600)  # gap from section label top to chart top

# Donut plot area AND legend are both left on PowerPoint's own fully
# automatic layout (no c:manualLayout at all) — manually pinning either
# one was tried across several iterations (fixed ring size, fixed legend
# position/height) and each version eventually caused some real chart to
# render wrong (ring inconsistent size, legend spread out, legend stuck at
# top-left, legend overlapping the ring for many-choice questions) in ways
# that couldn't be fully predicted without live rendering. PowerPoint's own
# auto-layout reliably avoids overlap and picks a sensible row/column
# arrangement for legendPos="b" on its own.
#
# DONUT_CENTER_Y_FRAC is therefore a plain heuristic (not backed by any
# guaranteed layout) for where the ring's visual center roughly falls, used
# only to position the Mean/NPS overlay text in _add_donut_center_text —
# it may be slightly off for questions with an unusually large or small
# number of legend items, since actual ring size now varies with how much
# room the auto-legend ends up needing.
DONUT_CENTER_Y_FRAC = 0.38

# Fraction of the stacked-chart's own allocated height reserved (at the top)
# for the 2-line ("7.2" / "-5.0") Mean/NPS label band above each column —
# the chart itself is shrunk by this much so the two never overlap (see
# _build_stacked).
STACK_MEAN_BAND_FRAC = 0.11

# ── Layout A: donut + 100%-stacked (slide2 in template) ──────────────────────
# Donut / breakdown split = 40% / 60% of the available content width.
DT_MARGIN_L = Emu(548640)
DT_GAP      = Emu(228600)
DT_CONTENT_W = int(SW) - 2 * int(DT_MARGIN_L)
_DT_AVAIL_W  = DT_CONTENT_W - int(DT_GAP)
DT_LEFT_W  = Emu(_DT_AVAIL_W * 40 // 100)
DT_RIGHT_W = Emu(_DT_AVAIL_W - int(DT_LEFT_W))
DT_RIGHT_L = Emu(int(DT_MARGIN_L) + int(DT_LEFT_W) + int(DT_GAP))

DT_TOTAL_L, DT_TOTAL_T = DT_MARGIN_L,  Emu(1143000)
DT_TOTAL_W              = DT_LEFT_W
DT_LEFT_L,  DT_LEFT_T   = DT_MARGIN_L,  Emu(1417320)
DT_LEFT_H               = Emu(4023360)
DT_RIGHT_LBL_L, DT_RIGHT_LBL_T = DT_RIGHT_L, Emu(1143000)
DT_RIGHT_LBL_W                  = DT_RIGHT_W
DT_RIGHT_T = Emu(1417320)
DT_RIGHT_H = Emu(4114800)

# ── Layout B: bar charts with N breakdown groups (slide1 in template) ─────────
# Total 40% / Sub-group 60% of content width
BR_TOTAL_L, BR_TOTAL_T = Emu(502920),  Emu(1097280)
BR_TOTAL_W              = Emu(4489704)
BR_LEFT_L,  BR_LEFT_T   = Emu(457200),  Emu(1371600)
BR_LEFT_W,  BR_LEFT_H   = Emu(4535424), Emu(4297680)
BR_RIGHT_LBL_L = Emu(5266944)   # x of right section labels
BR_RIGHT_LBL_W = Emu(6528816)   # width of right section labels
BR_RIGHT_L     = Emu(5221224)   # x of right charts
BR_RIGHT_W     = Emu(6574536)   # width of right charts
BR_RIGHT_FIRST_T = Emu(1097280)  # y of first right section label
BR_RIGHT_BOTTOM  = Emu(5715000)  # bottom of last right chart

# Max breakdown groups shown per slide (Layout B right side)
MAX_GROUPS_PER_SLIDE = 2

# Max stacked-chart breakdown COLUMNS shown per slide for donut_stacked
# (e.g. all banner groups combined can add up to 10+ columns — Area(4) +
# Gender(2) + Age(3) + HHI(4) = 13). Split across continuation slides,
# packing whole banner groups (never splitting one group's columns across
# two slides) — the Total donut is re-rendered identically on every
# continuation slide, only the stack's columns differ.
STACK_MAX_COLS_PER_SLIDE = 10

# ── Layout C: MA bar_horizontal — Total + up to 2 breakdowns, 3 equal columns ─
BH_GAP      = Emu(182880)   # gap between columns (0.2")
BH_MARGIN_L = Emu(457200)   # left margin (mirrors right margin)
BH_COL_W    = Emu((int(SW) - 2 * int(BH_MARGIN_L) - 2 * int(BH_GAP)) // 3)
BH_LBL_T    = Emu(1097280)
BH_CHART_T  = Emu(1371600)
BH_CHART_H  = Emu(4297680)

# Minimum Total-based percent for a choice to appear in an MA cross-tab chart —
# only applied when the question has at least MA_CROSSTAB_MIN_ITEMS choices;
# below that, cross-tab charts just hide 0% items like the Total chart does.
MA_CROSSTAB_MIN_PCT   = 0.10
MA_CROSSTAB_MIN_ITEMS = 6
MA_CROSSTAB_NOTE = (
    f"Note: cross-tab charts only show choices ≥{int(MA_CROSSTAB_MIN_PCT * 100)}% "
    f"of Total (applies to questions with {MA_CROSSTAB_MIN_ITEMS}+ choices)."
)


def _bh_col_l(i: int) -> Emu:
    """Left x-coordinate of the i-th (0-based) equal-width column in Layout C."""
    return Emu(int(BH_MARGIN_L) + i * (int(BH_COL_W) + int(BH_GAP)))


# ── Namespaces ────────────────────────────────────────────────────────────────
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_RID  = f"{{{_R_NS}}}id"


# ── Chart color palettes (extracted from temp.pptx chart XMLs) ───────────────

# Multi-series palette (breakdown col, stacked) — series 1→6
CHART_PALETTE = [
    "4472C4",  # blue
    "C0504D",  # red
    "9BBB59",  # green
    "95B3D7",  # light blue
    "A6A6A6",  # gray
    "D99694",  # pink/salmon
]

# Donut per-slice palette (matches chart7 in temp.pptx)
# 12 distinct colors — avoids repeats for most SA questions (donut/stack now
# used for all SA regardless of choice count, so >5-choice questions are
# common, e.g. occupation, province, shopping-behavior typology).
DONUT_PALETTE = [
    "4472C4",  # blue
    "95B3D7",  # light blue
    "A6A6A6",  # gray
    "D99694",  # pink/salmon
    "C0504D",  # red
    "9BBB59",  # green
    "8064A2",  # purple
    "F79646",  # orange
    "4BACC6",  # teal
    "808000",  # olive
    "17375E",  # dark navy
    "938953",  # brown/tan
]


# ── Font / text helpers ───────────────────────────────────────────────────────

def _set_run_arial(run) -> None:
    """Set font to Arial (latin + ea + cs) on a text run."""
    run.font.name = "Arial"
    rPr = run._r.get_or_add_rPr()
    for tag, pf, cs_val in [(qn("a:ea"), "34", "-122"), (qn("a:cs"), "34", "-120")]:
        el = rPr.find(tag)
        if el is None:
            el = etree.SubElement(rPr, tag)
        el.set("typeface", "Arial")
        el.set("pitchFamily", pf)
        el.set("charset", cs_val)


def _set_txbody_margins(tf, anchor: str = "ctr") -> None:
    tf.margin_left = 0
    tf.margin_top = 0
    tf.margin_right = 0
    tf.margin_bottom = 0
    tf.word_wrap = True
    # python-pptx's add_textbox() defaults to <a:spAutoFit/> (resize the
    # SHAPE to fit its text) — for short strings like "8.99"/"71.0" this
    # shrinks the box well below the width/position we explicitly computed,
    # which visually collapses a row of evenly-spread per-column labels
    # toward one edge instead of staying centered under their own column.
    # MSO_AUTO_SIZE.NONE (<a:noAutofit/>) keeps the exact box we asked for.
    tf.auto_size = MSO_AUTO_SIZE.NONE
    bodyPr = tf._txBody.find(qn("a:bodyPr"))
    if bodyPr is not None:
        bodyPr.set("anchor", anchor)


def _add_text(slide, left, top, width, height, text, *,
              font_size: int = 12, bold: bool = False,
              color: RGBColor | None = None,
              align=PP_ALIGN.LEFT, anchor: str = "ctr") -> None:
    txb = slide.shapes.add_textbox(left, top, width, height)
    tf = txb.text_frame
    _set_txbody_margins(tf, anchor)
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    _set_run_arial(run)


def _add_q_label(slide, left, top, width, height,
                 question: str, label: str, base) -> None:
    """Q-label: bold navy 'Qxx.' + gray description + gray bold (N=xx)."""
    txb = slide.shapes.add_textbox(left, top, width, height)
    tf = txb.text_frame
    _set_txbody_margins(tf, anchor="t")
    tf.word_wrap = True
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.LEFT

    def _run(text, bold=False, color=C_QDESC):
        r = para.add_run()
        r.text = text
        r.font.size = Pt(10)
        r.font.bold = bold
        r.font.color.rgb = color
        _set_run_arial(r)

    _run(f"{question}.  ", bold=True, color=C_NAVY)
    _run(label)
    _run(f"  (N={base})", bold=True, color=C_QBASE)


def _set_slide_note(slide, text: str) -> None:
    """Write text into the slide's speaker-notes pane ('Click to add notes')."""
    if not text:
        return
    slide.notes_slide.notes_text_frame.text = text


def _question_type_label(q: dict) -> str:
    """MA / SA-Ordinal / SA-Nominal, for the speaker-note type tag.

    Driven directly by chart_data's `scale_class` (Claude-classified in
    metadata.json, see CLAUDE.md Step 3c) — SA questions with
    `scale_class == "Ordinal"` are SA-Ordinal, everything else (including
    unclassified SA) is SA-Nominal."""
    answer_type = q.get("answer_type", "")
    if answer_type in ("MA", "Matrix_MA"):
        return "MA"
    if q.get("scale_class") == "Ordinal":
        return "SA-Ordinal"
    return "SA-Nominal"


def _add_rect(slide, left, top, width, height, fill: RGBColor) -> None:
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()


def _add_section_label(slide, left, top, width, text: str) -> None:
    _add_text(slide, left, top, width, SECTION_H, text,
              font_size=13, bold=True, color=C_NAVY,
              align=PP_ALIGN.CENTER, anchor="ctr")


# ── Label shortening ──────────────────────────────────────────────────────────

def _shorten_label(text: str) -> str:
    """Remove parenthetical clarifications and piping placeholders from labels, e.g.
    'By using ATM card (Internet banking)' → 'By using ATM card'
    'SHOW ANSWER OF Q13A : {QQ13A/792203/Selected}' → 'SHOW ANSWER OF Q13A :'"""
    result = re.sub(r"\s*\([^)]*\)", "", text)
    result = re.sub(r"\{[^{}]*\}", "", result).strip()
    result = re.sub(r"\s{2,}", " ", result)
    return re.sub(r"[\s:\-]+$", "", result)


def _shorten_labels(items: list, label_key: str = "label") -> list:
    """Shorten a list of choice labels via _shorten_label, but if two items
    collapse to the same shortened text — e.g. two age-banded variants of
    "Married, with children (youngest under 12)" / "...(youngest 12 or
    older)" both become "Married, with children" once the parenthetical
    (the only distinguishing part) is stripped — fall back to the original,
    unshortened text for just those colliding items so choices stay
    distinguishable in the chart/legend."""
    originals = [item[label_key] for item in items]
    shortened = [_shorten_label(t) for t in originals]
    counts: dict = {}
    for s in shortened:
        counts[s] = counts.get(s, 0) + 1
    return [
        orig if counts[short] > 1 else short
        for orig, short in zip(originals, shortened)
    ]


# ── Chart-template cloning ────────────────────────────────────────────────────

def _style_cat_axis(chartspace, *, font_pt: int = 7) -> None:
    """Reduce font size and enable word wrap on category axis labels. No rotation."""
    sz_val = str(font_pt * 100)  # DrawingML unit: 1/100 of a point

    for catAx in chartspace.iter(qn("c:catAx")):
        txPr = catAx.find(qn("c:txPr"))
        if txPr is None:
            txPr = etree.SubElement(catAx, qn("c:txPr"))

        bodyPr = txPr.find(qn("a:bodyPr"))
        if bodyPr is None:
            bodyPr = etree.SubElement(txPr, qn("a:bodyPr"))
        bodyPr.attrib.pop("rot", None)   # remove rotation if previously set
        bodyPr.set("wrap", "square")
        bodyPr.set("anchor", "ctr")

        if txPr.find(qn("a:lstStyle")) is None:
            etree.SubElement(txPr, qn("a:lstStyle"))

        p = txPr.find(qn("a:p"))
        if p is None:
            p = etree.SubElement(txPr, qn("a:p"))
        pPr = p.find(qn("a:pPr"))
        if pPr is None:
            pPr = etree.SubElement(p, qn("a:pPr"))
        defRPr = pPr.find(qn("a:defRPr"))
        if defRPr is None:
            defRPr = etree.SubElement(pPr, qn("a:defRPr"))
        defRPr.set("sz", sz_val)


def _style_legend(chartspace, *, font_pt: int = 7) -> None:
    """Force legend text (item/choice names) to a fixed size across every
    chart type, and force word-wrap so long entries stay fully readable
    within PowerPoint's fixed legend area instead of getting clipped."""
    sz_val = str(font_pt * 100)

    chart_el = chartspace.find(qn("c:chart"))
    if chart_el is None:
        return
    legend = chart_el.find(qn("c:legend"))
    if legend is None:
        return

    txPr = legend.find(qn("c:txPr"))
    if txPr is None:
        txPr = etree.SubElement(legend, qn("c:txPr"))

    bodyPr = txPr.find(qn("a:bodyPr"))
    if bodyPr is None:
        bodyPr = etree.SubElement(txPr, qn("a:bodyPr"))
    bodyPr.set("wrap", "square")
    if txPr.find(qn("a:lstStyle")) is None:
        etree.SubElement(txPr, qn("a:lstStyle"))

    p = txPr.find(qn("a:p"))
    if p is None:
        p = etree.SubElement(txPr, qn("a:p"))
    pPr = p.find(qn("a:pPr"))
    if pPr is None:
        pPr = etree.SubElement(p, qn("a:pPr"))
    defRPr = pPr.find(qn("a:defRPr"))
    if defRPr is None:
        defRPr = etree.SubElement(pPr, qn("a:defRPr"))
    defRPr.set("sz", sz_val)


def _force_pct_labels(chartspace) -> None:
    """Force every data-label number format to 0"%" (values injected as 0-100)."""
    for dl in chartspace.iter(qn("c:dLbls")):
        nf = dl.find(qn("c:numFmt"))
        if nf is None:
            nf = etree.Element(qn("c:numFmt"))
            dl.insert(0, nf)
        nf.set("formatCode", '0"%"')
        nf.set("sourceLinked", "0")



# CT_DLbls / CT_DLbl: elements that must come AFTER c:txPr in document order —
# used to find the correct insertion point when txPr is missing (a plain
# append would violate the schema if any of these are already present).
_DLBL_AFTER_TXPR = (
    "c:dLblPos", "c:showLegendKey", "c:showVal", "c:showCatName",
    "c:showSerName", "c:showPercent", "c:showBubbleSize", "c:separator",
    "c:showLeaderLines", "c:leaderLines", "c:extLst",
)


def _insert_txpr(dl_container):
    """Create a <c:txPr> on a <c:dLbls>/<c:dLbl> element at the schema-correct
    position (before dLblPos/showVal/etc., after idx/numFmt/spPr/any nested
    dLbl), rather than blindly appending at the end."""
    insert_at = len(dl_container)
    for i, child in enumerate(dl_container):
        if etree.QName(child).text in {qn(t) for t in _DLBL_AFTER_TXPR}:
            insert_at = i
            break
    txPr = etree.Element(qn("c:txPr"))
    etree.SubElement(txPr, qn("a:bodyPr"))
    etree.SubElement(txPr, qn("a:lstStyle"))
    dl_container.insert(insert_at, txPr)
    return txPr


def _style_data_labels(chartspace, *, font_pt: int = 6) -> None:
    """Force the percent value shown on every data label (donut slice, bar/
    column end, stacked segment — series-level default AND any per-point
    <c:dLbl> override) to a fixed font size, across every chart type."""
    sz_val = str(font_pt * 100)
    for dl_container in list(chartspace.iter(qn("c:dLbls"))) + list(chartspace.iter(qn("c:dLbl"))):
        txPr = dl_container.find(qn("c:txPr"))
        if txPr is None:
            txPr = _insert_txpr(dl_container)
        p = txPr.find(qn("a:p"))
        if p is None:
            p = etree.SubElement(txPr, qn("a:p"))
        pPr = p.find(qn("a:pPr"))
        if pPr is None:
            pPr = etree.SubElement(p, qn("a:pPr"))
        defRPr = pPr.find(qn("a:defRPr"))
        if defRPr is None:
            defRPr = etree.SubElement(pPr, qn("a:defRPr"))
        defRPr.set("sz", sz_val)


def _hide_zero_value_labels(chartspace) -> None:
    """Delete the data label (on-chart % text / slice) for any point whose
    value is 0, while leaving it in the category/legend cache untouched — so
    the item still appears in the legend, just without a visible 0% slice."""
    chart_el = chartspace.find(qn("c:chart"))
    if chart_el is None:
        return
    plot_area = chart_el.find(qn("c:plotArea"))
    if plot_area is None:
        return
    for ser in plot_area.iter(qn("c:ser")):
        val_el = ser.find(qn("c:val"))
        if val_el is None:
            continue
        numRef = val_el.find(qn("c:numRef"))
        if numRef is None:
            continue
        cache = numRef.find(qn("c:numCache"))
        if cache is None:
            continue
        zero_idxs = sorted(
            int(pt.get("idx")) for pt in cache.findall(qn("c:pt"))
            if float((pt.find(qn("c:v")).text or "0")) == 0
        )
        if not zero_idxs:
            continue
        dLbls = ser.find(qn("c:dLbls"))
        if dLbls is None:
            continue

        # A dLbl per idx must be unique (CT_DLbl: idx, then either <delete> or
        # layout/format children) — reuse the template's existing override for
        # that idx instead of appending a second, conflicting one.
        existing_by_idx = {}
        for dLbl in dLbls.findall(qn("c:dLbl")):
            idx_el = dLbl.find(qn("c:idx"))
            if idx_el is not None:
                existing_by_idx[int(idx_el.get("val"))] = dLbl

        insert_at = 0
        for idx in zero_idxs:
            dLbl = existing_by_idx.get(idx)
            if dLbl is not None:
                for child in list(dLbl):
                    if child.tag != qn("c:idx"):
                        dLbl.remove(child)
                etree.SubElement(dLbl, qn("c:delete")).set("val", "1")
            else:
                dLbl = etree.Element(qn("c:dLbl"))
                etree.SubElement(dLbl, qn("c:idx")).set("val", str(idx))
                etree.SubElement(dLbl, qn("c:delete")).set("val", "1")
                dLbls.insert(insert_at, dLbl)
                insert_at += 1


def _remove_legend(chartspace) -> None:
    chart_el = chartspace.find(qn("c:chart"))
    if chart_el is None:
        return
    legend = chart_el.find(qn("c:legend"))
    if legend is not None:
        chart_el.remove(legend)


def _ensure_legend(chartspace, *, pos: str = "b") -> None:
    """Add a <c:legend> if the template doesn't have one (the "bar" template
    has none, since it's normally used single-series for the Total chart —
    but a multi-series breakdown bar chart needs one to tell series apart).
    No-op if a legend already exists."""
    chart_el = chartspace.find(qn("c:chart"))
    if chart_el is None:
        return
    if chart_el.find(qn("c:legend")) is not None:
        return
    legend = etree.Element(qn("c:legend"))
    etree.SubElement(legend, qn("c:legendPos")).set("val", pos)
    etree.SubElement(legend, qn("c:overlay")).set("val", "0")
    txPr = etree.SubElement(legend, qn("c:txPr"))
    etree.SubElement(txPr, qn("a:bodyPr"))
    etree.SubElement(txPr, qn("a:lstStyle"))
    p = etree.SubElement(txPr, qn("a:p"))
    pPr = etree.SubElement(p, qn("a:pPr"))
    defRPr = etree.SubElement(pPr, qn("a:defRPr"))
    defRPr.set("sz", "900")
    etree.SubElement(defRPr, qn("a:latin")).set("typeface", "Arial")
    etree.SubElement(defRPr, qn("a:cs")).set("typeface", "Arial")
    etree.SubElement(p, qn("a:endParaRPr")).set("lang", "en-US")

    # CT_Chart schema order: ... plotArea, legend?, plotVisOnly?, ...
    plot_area = chart_el.find(qn("c:plotArea"))
    insert_at = list(chart_el).index(plot_area) + 1 if plot_area is not None else len(chart_el)
    chart_el.insert(insert_at, legend)


def _match_series_count(chartspace, n: int) -> None:
    """Make the template's plot have exactly n <c:ser> elements before
    replace_data() runs, by cloning/trimming the template's own series.

    Root-cause fix for a PowerPoint quirk (confirmed via real PowerPoint
    rendering): when replace_data() itself has to add/remove series because
    the new data's series count differs from the template's fixed sample
    count, the chart's LEGEND can render in a different order than the
    series were added in — inconsistently, depending on the count. Matching
    the series count ourselves beforehand means replace_data() only ever
    updates existing series content (never adds/removes any), which renders
    with the legend in the expected add order every time."""
    chart_el = chartspace.find(qn("c:chart"))
    if chart_el is None:
        return
    plot_area = chart_el.find(qn("c:plotArea"))
    if plot_area is None:
        return
    parent = next((child for child in plot_area
                    if child.find(qn("c:ser")) is not None), None)
    if parent is None:
        return
    sers = parent.findall(qn("c:ser"))
    if not sers or n <= 0:
        return
    if len(sers) < n:
        template_ser = sers[-1]
        while len(parent.findall(qn("c:ser"))) < n:
            parent.append(copy.deepcopy(template_ser))
    elif len(sers) > n:
        for extra in parent.findall(qn("c:ser"))[n:]:
            parent.remove(extra)
    for i, ser in enumerate(parent.findall(qn("c:ser"))):
        idx_el = ser.find(qn("c:idx"))
        order_el = ser.find(qn("c:order"))
        if idx_el is not None:
            idx_el.set("val", str(i))
        if order_el is not None:
            order_el.set("val", str(i))


def _match_point_dlbl_count(chartspace) -> None:
    """Extend/trim a single-series chart's per-point <c:dLbl> overrides (e.g.
    donut) to match the actual point count, cloning the last override's
    formatting for any extra points.

    Root-cause fix (confirmed via real PowerPoint rendering): the donut
    template only ships per-point dLbl overrides for its original 5 sample
    slices. Points beyond that (e.g. a 7-choice SA question, now that all SA
    questions use the donut layout) have no override and fall back to the
    series-level defaults — which rendered wrong (showed "0%" instead of the
    real percentage) in real PowerPoint. Giving every point its own override,
    cloned from the same template styling, fixes this for any point count."""
    chart_el = chartspace.find(qn("c:chart"))
    if chart_el is None:
        return
    plot_area = chart_el.find(qn("c:plotArea"))
    if plot_area is None:
        return
    all_sers = [s for child in plot_area for s in child.findall(qn("c:ser"))]
    if len(all_sers) != 1:
        return
    ser = all_sers[0]

    val_el = ser.find(qn("c:val"))
    n = 0
    numRef = val_el.find(qn("c:numRef")) if val_el is not None else None
    if numRef is not None:
        cache = numRef.find(qn("c:numCache"))
        if cache is not None:
            n = len(cache.findall(qn("c:pt")))
    if n == 0:
        return

    dLbls = ser.find(qn("c:dLbls"))
    if dLbls is None:
        return
    dLbl_els = dLbls.findall(qn("c:dLbl"))
    if not dLbl_els:
        return

    if len(dLbl_els) < n:
        template_dLbl = dLbl_els[-1]
        insert_pos = list(dLbls).index(template_dLbl) + 1
        for i in range(len(dLbl_els), n):
            new_dLbl = copy.deepcopy(template_dLbl)
            new_dLbl.find(qn("c:idx")).set("val", str(i))
            dLbls.insert(insert_pos, new_dLbl)
            insert_pos += 1
    elif len(dLbl_els) > n:
        for extra in dLbl_els[n:]:
            dLbls.remove(extra)


def _spPr_solid(hex_color: str):
    """<c:spPr> with solid fill and no border line."""
    spPr = etree.Element(qn("c:spPr"))
    solidFill = etree.SubElement(spPr, qn("a:solidFill"))
    etree.SubElement(solidFill, qn("a:srgbClr")).set("val", hex_color)
    etree.SubElement(etree.SubElement(spPr, qn("a:ln")), qn("a:noFill"))
    return spPr


def _set_ser_spPr(ser, hex_color: str) -> None:
    """Replace or insert <c:spPr> on a series element."""
    existing = ser.find(qn("c:spPr"))
    if existing is not None:
        ser.remove(existing)
    order_el = ser.find(qn("c:order"))
    if order_el is not None:
        ser.insert(list(ser).index(order_el) + 1, _spPr_solid(hex_color))
    else:
        ser.append(_spPr_solid(hex_color))


def _apply_palette(chartspace, palette: list, *, per_point: bool = False,
                   series_colors: list | None = None) -> None:
    """Apply palette colors to chart series.

    per_point=False (bar/col):  single series → 1 color on the series itself.
    per_point=True  (donut):    single series → per-slice dPt colors.
    multi-series (any):         one color per series via spPr — uses
                                 `series_colors[i]` if given (so callers can
                                 keep colors consistent with a paired chart),
                                 otherwise falls back to `palette[i % len]`.
    """
    chart_el = chartspace.find(qn("c:chart"))
    if chart_el is None:
        return
    plot_area = chart_el.find(qn("c:plotArea"))
    if plot_area is None:
        return

    all_sers = [s for child in plot_area for s in child.findall(qn("c:ser"))]
    if not all_sers:
        return

    if len(all_sers) > 1:
        for i, ser in enumerate(all_sers):
            color = series_colors[i] if series_colors else palette[i % len(palette)]
            _set_ser_spPr(ser, color)
        return

    ser = all_sers[0]
    if not per_point:
        # Bar/col single series: uniform color
        _set_ser_spPr(ser, series_colors[0] if series_colors else palette[0])
    else:
        # Donut: color each slice individually via dPt
        for dPt in ser.findall(qn("c:dPt")):
            ser.remove(dPt)
        val_el = ser.find(qn("c:val"))
        n = 0
        if val_el is not None:
            numRef = val_el.find(qn("c:numRef"))
            if numRef is not None:
                cache = numRef.find(qn("c:numCache"))
                if cache is not None:
                    n = len(cache.findall(qn("c:pt")))
        if n == 0:
            return
        children = list(ser)
        insert_pos = next(
            (i for i, c in enumerate(children)
             if c.tag in (qn("c:cat"), qn("c:val"))),
            len(children),
        )
        for i in range(n):
            dPt = etree.Element(qn("c:dPt"))
            etree.SubElement(dPt, qn("c:idx")).set("val", str(i))
            dPt.append(_spPr_solid(palette[i % len(palette)]))
            ser.insert(insert_pos + i, dPt)


def _clone_chart(slide, tmpl_chartspace, pptx_type, l, t, w, h, chart_data,
                 *, drop_legend: bool = False, per_point: bool = False,
                 palette: list | None = None,
                 cat_font_pt: int = 7, legend_n: int | None = None,
                 hide_zero_labels: bool = False, series_colors: list | None = None):
    """Add a chart, replace its XML with a deep copy of the template chartSpace,
    then inject `chart_data` via replace_data() so template styling is kept.

    Item text (category axis + legend) and data-label percent text are
    always forced to a fixed size (7pt / 6pt) for every chart type — see
    _style_cat_axis / _style_legend / _style_data_labels. Both are no-ops
    when the chart has no category axis (donut) or no legend, so calling
    them unconditionally is safe."""
    gf = slide.shapes.add_chart(pptx_type, l, t, w, h, chart_data)
    cp = gf.chart_part

    new_ext = cp._element.find(qn("c:externalData"))
    new_rid = new_ext.get(_RID) if new_ext is not None else None

    cs = copy.deepcopy(tmpl_chartspace)
    t_ext = cs.find(qn("c:externalData"))
    if t_ext is not None and new_rid is not None:
        t_ext.set(_RID, new_rid)

    if drop_legend:
        _remove_legend(cs)
    _match_series_count(cs, len(chart_data))

    cp._element = cs
    cp.__dict__.pop("chart", None)
    cp.__dict__.pop("chart_workbook", None)

    cp.chart.replace_data(chart_data)
    _force_pct_labels(cp._element)
    _style_cat_axis(cp._element, font_pt=cat_font_pt)
    if legend_n is not None:
        _ensure_legend(cp._element)
    _style_legend(cp._element)
    if per_point:
        _match_point_dlbl_count(cp._element)
    if hide_zero_labels:
        _hide_zero_value_labels(cp._element)
    _style_data_labels(cp._element)
    _apply_palette(cp._element, palette or CHART_PALETTE, per_point=per_point,
                   series_colors=series_colors)
    return cp.chart


def _v100(percents: dict, codes: list) -> list:
    """Fractions (0-1) to percentages (0-100) for the given choice codes."""
    return [round(percents.get(c, 0.0) * 100.0, 4) for c in codes]


def _mean_nps_lines(data: dict) -> list[str]:
    """['Mean: 7.2'] / ['Mean: 7.2', 'NPS: -5.0'] for a total/breakdown-column
    dict that carries Step 3c's `mean`/`nps` stat values — Mean and NPS are
    always separate lines (never joined on one line), and this returns []
    when the question has no `mean` (e.g. Nominal or MA), so it's always
    safe to call unconditionally."""
    mean_v = data.get("mean")
    if mean_v is None:
        return []
    lines = [f"Mean: {mean_v:g}"]
    nps_v = data.get("nps")
    if nps_v is not None:
        lines.append(f"NPS: {nps_v:.1f}")
    return lines


def _mean_nps_compact(data: dict) -> list[str]:
    """['9.24'] / ['9.24', '82.8'] — bare numbers, one per line (Mean then
    NPS), used for the stack chart's per-column top labels. A single
    "Mean" / "NPS" caption to the left labels the two rows once (see
    _build_stacked), rather than repeating "Mean:"/"NPS:" prefixes on
    every column."""
    mean_v = data.get("mean")
    if mean_v is None:
        return []
    lines = [f"{mean_v:g}"]
    nps_v = data.get("nps")
    if nps_v is not None:
        lines.append(f"{nps_v:.1f}")
    return lines


def _col_label_with_base(col: dict) -> str:
    """Breakdown column label + its own base, e.g. 'Male (N=120)'.

    Mean/NPS (Step 3c Ordinal questions) is NOT appended here — it's shown
    as a data label sitting on top of that column's stacked bar instead
    (see _set_stack_top_labels), so the x-axis stays a plain category name."""
    return f"{col.get('label', '')} (N={col.get('base', 0)})"


# Meta codes that are never part of an ordinal scale itself, even when they
# appear alongside one — per this project's own FT-coding convention (see
# CLAUDE.md Workflow C: 98 = "Others", 99 = "No answer"). Sorting these
# purely by numeric code value would put them at the very top of a 1-5
# scale (since 98/99 > 5), which is never the intended meaning.
_ORDINAL_META_CODES = {98, 99}


def _code_int(choice: dict) -> int | None:
    """Parse a single choice's code as int, or None if it isn't numeric."""
    try:
        return int(choice["code"])
    except (ValueError, TypeError):
        return None


def _sort_scale_desc(q: dict, choices: list) -> list:
    """SA-Ordinal questions (`scale_class == "Ordinal"`, Claude-classified in
    metadata.json — see CLAUDE.md Step 3c): sort choices highest scale point
    to lowest. SA-Nominal questions are returned unchanged.

    Which code is "highest" isn't always the max code value — e.g. a
    frequency scale coded 1=Almost every day … 7=Less than once a month has
    code 1 at the high-intensity end. `scale_high_code` (also
    Claude-classified, see CLAUDE.md Step 3c) names the code at the high end
    explicitly; sort direction is derived from whether that's the min or max
    code among the real scale choices (excluding meta codes below).
    Defaults to "max code = high end" (descending by code) when
    `scale_high_code` isn't set.

    Choices with a non-numeric code, or a meta code (_ORDINAL_META_CODES),
    are excluded from the scale sort and always appended at the end (in
    their original relative order) — sorting them in with the real scale
    values would place e.g. "Others"/"No answer" at the top purely because
    98/99 is numerically larger than the scale's real range."""
    if q.get("scale_class") != "Ordinal":
        return choices

    scale_choices, meta_choices = [], []
    for c in choices:
        code = _code_int(c)
        if code is None or code in _ORDINAL_META_CODES:
            meta_choices.append(c)
        else:
            scale_choices.append(c)
    if not scale_choices:
        return choices

    codes = [_code_int(c) for c in scale_choices]
    high_code = q.get("scale_high_code")
    reverse = True
    if high_code is not None:
        try:
            reverse = int(high_code) != min(codes)
        except (ValueError, TypeError):
            pass
    scale_choices.sort(key=_code_int, reverse=reverse)
    return scale_choices + meta_choices


def _order_sa_choices(q: dict, choices: list) -> list:
    """Order SA (donut_stacked) choices for both the donut and its paired
    stacked breakdown chart, so the two always agree:

    - NPS questions (`is_nps`, see CLAUDE.md Step 3c): never re-sorted —
      chart_data_exporter already emits exactly 3 choices (Promoters,
      Passives, Detractors) in that fixed order, keyed by synthetic codes
      "0"/"1"/"2". Sorting by scale value here would be meaningless (and
      actively wrong) since those aren't real scale codes.
    - SA-Ordinal: sort by scale order descending (see _sort_scale_desc).
    - SA-Nominal: sort by Total percent descending, tie-broken by ascending
      code (stable/deterministic tie-break, rather than silently depending
      on whatever order the choices happened to arrive in — which mattered
      in practice for questions with several 0%/tied choices)."""
    if q.get("is_nps"):
        return choices
    if q.get("scale_class") == "Ordinal":
        return _sort_scale_desc(q, choices)
    total_pct = q.get("total", {}).get("percents", {})

    def _key(c):
        code = _code_int(c)
        return (-total_pct.get(c["code"], 0), code if code is not None else 0)

    return sorted(choices, key=_key)


# ── Per-chart builders (each clones the matching template chart) ──────────────

def _build_total_bar(slide, q, tmpl, l, t, w, h) -> None:
    """Horizontal bar, single series, sorted ascending (largest at top). Hides 0% items."""
    choices = q["choices"]
    pcts = q.get("total", {}).get("percents", {})
    choices = [c for c in choices if pcts.get(c["code"], 0) > 0]
    if not choices:
        return
    codes = [c["code"] for c in choices]
    labels = _shorten_labels(choices)
    order = sorted(range(len(codes)), key=lambda i: pcts.get(codes[i], 0))
    labels = [labels[i] for i in order]
    codes  = [codes[i]  for i in order]
    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series("%", _v100(pcts, codes))
    _clone_chart(slide, tmpl["bar"], XL_CHART_TYPE.BAR_CLUSTERED, l, t, w, h, cd)


def _build_total_col(slide, q, tmpl, l, t, w, h) -> None:
    """Vertical column, single series. Hides 0% items. SA-Likert questions
    (scale detected) sort descending by scale point; everything else keeps
    natural order."""
    choices = q["choices"]
    pcts = q.get("total", {}).get("percents", {})
    choices = [c for c in choices if pcts.get(c["code"], 0) > 0]
    if not choices:
        return
    choices = _sort_scale_desc(q, choices)
    codes  = [c["code"] for c in choices]
    labels = _shorten_labels(choices)
    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series("Total", _v100(pcts, codes))
    _clone_chart(slide, tmpl["col"], XL_CHART_TYPE.COLUMN_CLUSTERED, l, t, w, h, cd,
                 drop_legend=True)


def _build_breakdown_col(slide, q, group, tmpl, l, t, w, h) -> None:
    """Vertical clustered columns — one series per breakdown column. Hides 0%
    items. SA-Likert questions (scale detected) sort descending by scale
    point, matching the Total chart; everything else keeps natural order."""
    choices = q["choices"]
    cols = group["columns"]
    choices = [c for c in choices
               if any(col.get("percents", {}).get(c["code"], 0) > 0 for col in cols)]
    if not choices:
        return
    choices = _sort_scale_desc(q, choices)
    codes  = [c["code"] for c in choices]
    labels = _shorten_labels(choices)
    cd = CategoryChartData()
    cd.categories = labels
    for col in cols:
        cd.add_series(col["label"], _v100(col.get("percents", {}), codes))
    _clone_chart(slide, tmpl["col"], XL_CHART_TYPE.COLUMN_CLUSTERED, l, t, w, h, cd)


def _build_breakdown_bar(slide, q, group, tmpl, l, t, w, h) -> None:
    """Horizontal clustered bars — one series per breakdown column, ordered like
    the Total chart (largest at top). All MA bar charts hide 0% items; for
    questions with >= MA_CROSSTAB_MIN_ITEMS choices, the MA_CROSSTAB_MIN_PCT
    cutoff is applied on top of that."""
    choices = q["choices"]
    cols = group["columns"]
    total_pct = q.get("total", {}).get("percents", {})
    if len(choices) >= MA_CROSSTAB_MIN_ITEMS:
        choices = [c for c in choices if total_pct.get(c["code"], 0) >= MA_CROSSTAB_MIN_PCT]
    else:
        choices = [c for c in choices if total_pct.get(c["code"], 0) > 0]
    if not choices:
        return
    codes  = [c["code"] for c in choices]
    labels = _shorten_labels(choices)
    order = sorted(range(len(codes)), key=lambda i: total_pct.get(codes[i], 0))
    codes  = [codes[i]  for i in order]
    labels = [labels[i] for i in order]
    cd = CategoryChartData()
    cd.categories = labels
    for col in cols:
        cd.add_series(col["label"], _v100(col.get("percents", {}), codes))
    _clone_chart(slide, tmpl["bar"], XL_CHART_TYPE.BAR_CLUSTERED, l, t, w, h, cd,
                 legend_n=len(cols) if len(cols) > 1 else None)


def _build_breakdown_stacked(slide, q, group, tmpl, l, t, w, h, *,
                             horizontal: bool = False) -> None:
    """100%-stacked breakdown chart for questions with many choices (> STACKED_THRESHOLD).
    Categories = subgroup columns (e.g. Male / Female).
    Series     = choices (items), 0% filtered.
    horizontal=True  → BAR_STACKED_100  (for MA bar_horizontal questions)
    horizontal=False → COLUMN_STACKED_100
    """
    choices = q["choices"]
    cols = group["columns"]
    choices = [c for c in choices
               if any(col.get("percents", {}).get(c["code"], 0) > 0 for col in cols)]
    if not choices:
        return
    cd = CategoryChartData()
    cd.categories = [_col_label_with_base(col) for col in cols]
    choice_labels = _shorten_labels(choices)
    for choice, choice_label in zip(choices, choice_labels):
        cd.add_series(
            choice_label,
            [round(col.get("percents", {}).get(choice["code"], 0.0) * 100.0, 4)
             for col in cols],
        )
    if horizontal:
        _clone_chart(slide, tmpl["bar"], XL_CHART_TYPE.BAR_STACKED_100,
                     l, t, w, h, cd)
    else:
        _clone_chart(slide, tmpl["stacked"], XL_CHART_TYPE.COLUMN_STACKED_100,
                     l, t, w, h, cd)


def _add_multiline_text(slide, left, top, width, height, lines: list[str], *,
                        font_size: int = 9, bold: bool = True,
                        color: RGBColor | None = None) -> None:
    """Small centered text box with each string in `lines` as its own
    paragraph (never joined with a separator onto one line)."""
    txb = slide.shapes.add_textbox(left, top, width, height)
    tf = txb.text_frame
    _set_txbody_margins(tf, anchor="ctr")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = line
        r.font.size = Pt(font_size)
        r.font.bold = bold
        if color:
            r.font.color.rgb = color
        _set_run_arial(r)


def _add_donut_center_text(slide, l, t, w, h, lines: list[str]) -> None:
    """Small Mean/NPS text box sitting inside the donut's hole (holeSize=55%
    of radius, see donut.xml). The chart's own legend is at the bottom
    (legendPos="b"), so the ring's visual center sits somewhat above the
    shape's full vertical midpoint — approximated here since exact
    placement depends on PowerPoint's own auto plot-area layout (no manual
    layout is set in the template), which can't be measured without
    rendering; nudge DONUT_CENTER_Y_FRAC if it's off once seen rendered."""
    box_w = Emu(int(w) * 2 // 5)
    box_h = Emu(int(h) // 6)
    box_l = Emu(int(l) + (int(w) - int(box_w)) // 2)
    box_t = Emu(int(t) + int(int(h) * DONUT_CENTER_Y_FRAC) - int(box_h) // 2)
    _add_multiline_text(slide, box_l, box_t, box_w, box_h, lines,
                        font_size=9, bold=True, color=C_NAVY)


def _build_donut(slide, q, tmpl, l, t, w, h) -> None:
    """Donut chart for SA≤5. Every choice (incl. 0%) stays in the legend, but
    0%-value slices/labels are hidden on the chart itself. Order: scale
    questions descending by scale point, non-scale by Total percent
    descending (see _order_sa_choices)."""
    choices = q["choices"]
    pcts = q.get("total", {}).get("percents", {})
    if not choices:
        return
    choices = _order_sa_choices(q, choices)
    codes  = [c["code"] for c in choices]
    labels = _shorten_labels(choices)
    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series("Total", _v100(pcts, codes))
    _clone_chart(slide, tmpl["donut"], XL_CHART_TYPE.DOUGHNUT, l, t, w, h, cd,
                 per_point=True, palette=DONUT_PALETTE, legend_n=len(labels),
                 hide_zero_labels=True)
    center_lines = _mean_nps_lines(q.get("total", {}))
    if center_lines:
        _add_donut_center_text(slide, l, t, w, h, center_lines)


def _add_stack_top_labels(slide, l, t, w, band_h, lines_per_col: list[list[str]]) -> None:
    """A small 2-line ("7.2" / "-5.0") Mean/NPS textbox per column, each one
    centered on that column's own horizontal slot (w / n, matching how
    PowerPoint centers a category's bar within its own equal-width slot —
    same assumption category axes use by default), inside the reserved top
    band (l, t, w, band_h). The caller has already shrunk the actual chart
    to start below this band, so there's no dependency on estimating where
    the real chart's plot area happens to render; the two literally cannot
    overlap.

    A single 2-line "Mean" / "NPS" caption is added once, to the left of
    the first column (in the gap before the chart's own left edge), lined
    up row-for-row with the per-column values — rather than repeating
    "Mean:"/"NPS:" labels on every column."""
    n = len(lines_per_col)
    if n == 0:
        return
    col_w = int(w) // n
    box_w = Emu(int(col_w * 0.92))
    box_h = Emu(int(band_h))

    # Only label "NPS" if at least one column actually has an nps value —
    # a Mean-only Ordinal question (no NPS stat) should show just "Mean".
    has_nps = any(len(lines) > 1 for lines in lines_per_col)
    caption = ["Mean", "NPS"] if has_nps else ["Mean"]
    caption_l = Emu(max(0, int(l) - int(box_w)))
    _add_multiline_text(slide, caption_l, Emu(int(t)), box_w, box_h,
                        caption, font_size=7, bold=True, color=C_QBASE)

    for i, lines in enumerate(lines_per_col):
        if not lines:
            continue
        col_l = int(l) + i * col_w
        box_l = Emu(col_l + (col_w - int(box_w)) // 2)
        _add_multiline_text(slide, box_l, Emu(int(t)), box_w, box_h, lines,
                            font_size=7, bold=True, color=C_NAVY)


def _build_stacked(slide, q, breakdowns, tmpl, l, t, w, h) -> None:
    """100%-stacked columns — series per choice, using the same DONUT_PALETTE
    colors per choice so the two charts agree. Shows every choice (incl.
    those that are 0% everywhere) so the legend always matches the donut's
    full list — donut's post-sort order (_order_sa_choices) is the source of
    truth for both the stack's item order (within a column) and its legend
    order.

    Series are always added in the same order as the donut's sorted choices
    (SA-Nominal / SA-Ordinal / Matrix-SA alike) — no reversal. PowerPoint's
    legend for THIS chart can render in reverse of series-add order when
    labels are long enough to force a single-column legend layout (observed
    on long-label questions like S10/A5/B6) — accepted as a known edge case
    rather than compensated for, since a label-length-based reversal rule
    was tried and rejected: it fixed those long-label cases but flipped the
    order on other (shorter-label) Nominal questions that were already
    correct without any reversal."""
    choices = q["choices"]
    all_cols, xlabels = [], []
    for bd in breakdowns:
        for col in bd["columns"]:
            all_cols.append(col)
            xlabels.append(_col_label_with_base(col))
    if not all_cols or not choices:
        return
    choices = _order_sa_choices(q, choices)
    color_map = {c["code"]: DONUT_PALETTE[i % len(DONUT_PALETTE)]
                 for i, c in enumerate(choices)}
    shortened = {c["code"]: lbl for c, lbl in zip(choices, _shorten_labels(choices))}
    cd = CategoryChartData()
    cd.categories = xlabels
    series_colors = []
    for choice in choices:
        cd.add_series(
            shortened[choice["code"]],
            [round(col.get("percents", {}).get(choice["code"], 0.0) * 100.0, 4)
             for col in all_cols],
        )
        series_colors.append(color_map[choice["code"]])

    # Mean/NPS (Step 3c Ordinal questions) sits above each column, in a
    # reserved band carved out of the TOP of this function's own (l, t, w, h)
    # allocation — the chart itself is shrunk/shifted down to make room, so
    # the label band and the chart body never overlap.
    mean_lines = [_mean_nps_compact(col) for col in all_cols]
    if any(mean_lines):
        chart_t = Emu(int(t) + int(int(h) * STACK_MEAN_BAND_FRAC))
        chart_h = Emu(int(h) - int(int(h) * STACK_MEAN_BAND_FRAC))
        _add_stack_top_labels(slide, l, t, w, int(int(h) * STACK_MEAN_BAND_FRAC), mean_lines)
    else:
        chart_t, chart_h = t, h

    _clone_chart(slide, tmpl["stacked"], XL_CHART_TYPE.COLUMN_STACKED_100, l, chart_t, w, chart_h, cd,
                series_colors=series_colors,
                hide_zero_labels=True, legend_n=len(choices))


# ── Slide builder ─────────────────────────────────────────────────────────────

def _build_slide(prs, layout, q, page_num, tmpl, *, title_suffix: str = "") -> None:
    slide = prs.slides.add_slide(layout)
    chart_type = q.get("chart_type", "bar_vertical")
    breakdowns = q.get("breakdowns", [])
    n_groups   = len(breakdowns)
    note_lines = [f"Question type: {_question_type_label(q)}"]

    q_label    = _shorten_label(q.get("label", ""))
    title_text = f"{q_label} {title_suffix}".strip() if title_suffix else q_label

    _add_rect(slide, FOOTER_L, FOOTER_T, FOOTER_W, FOOTER_H, C_ORANGE)
    title_font = 18 if len(title_text) > 120 else 22 if len(title_text) > 80 else 30
    _add_text(slide, TITLE_L, TITLE_T, TITLE_W, TITLE_H,
              title_text, font_size=title_font, bold=True, color=C_NAVY)

    if chart_type == "donut_stacked":
        _add_section_label(slide, DT_TOTAL_L, DT_TOTAL_T, DT_TOTAL_W, "Total")
        _build_donut(slide, q, tmpl, DT_LEFT_L, DT_LEFT_T, DT_LEFT_W, DT_LEFT_H)
        if breakdowns:
            combined = " / ".join(bd["group_label"] for bd in breakdowns)
            _add_section_label(slide, DT_RIGHT_LBL_L, DT_RIGHT_LBL_T,
                               DT_RIGHT_LBL_W, combined)
            _build_stacked(slide, q, breakdowns, tmpl,
                           DT_RIGHT_L, DT_RIGHT_T, DT_RIGHT_W, DT_RIGHT_H)
    elif chart_type == "bar_horizontal":
        # MA: Total + up to MAX_GROUPS_PER_SLIDE breakdowns, 3 equal-width columns.
        _add_section_label(slide, _bh_col_l(0), BH_LBL_T, BH_COL_W, "Total")
        _build_total_bar(slide, q, tmpl, _bh_col_l(0), BH_CHART_T, BH_COL_W, BH_CHART_H)
        for gi, bd in enumerate(breakdowns[:MAX_GROUPS_PER_SLIDE]):
            col_l = _bh_col_l(gi + 1)
            _add_section_label(slide, col_l, BH_LBL_T, BH_COL_W, bd["group_label"])
            _build_breakdown_bar(slide, q, bd, tmpl,
                                 col_l, BH_CHART_T, BH_COL_W, BH_CHART_H)
        if breakdowns and len(q.get("choices", [])) >= MA_CROSSTAB_MIN_ITEMS:
            note_lines.append(MA_CROSSTAB_NOTE)

    else:
        _add_section_label(slide, BR_TOTAL_L, BR_TOTAL_T, BR_TOTAL_W, "Total")
        _build_total_col(slide, q, tmpl, BR_LEFT_L, BR_LEFT_T, BR_LEFT_W, BR_LEFT_H)

        if breakdowns:
            total_span = int(BR_RIGHT_BOTTOM - BR_RIGHT_FIRST_T)
            per_group  = total_span // max(n_groups, 1)
            chart_h    = per_group - int(SECTION_TO_CHART)
            for gi, bd in enumerate(breakdowns):
                lbl_t   = int(BR_RIGHT_FIRST_T) + gi * per_group
                chart_t = lbl_t + int(SECTION_TO_CHART)
                _add_section_label(slide, BR_RIGHT_LBL_L, Emu(lbl_t),
                                   BR_RIGHT_LBL_W, bd["group_label"])
                _build_breakdown_col(slide, q, bd, tmpl,
                                         BR_RIGHT_L, Emu(chart_t),
                                         BR_RIGHT_W, Emu(chart_h))

    q_base = q.get("total", {}).get("base", "XX")
    _add_q_label(slide, Q_LABEL_L, Q_LABEL_T, Q_LABEL_W, Q_LABEL_H,
                 q.get("question", ""), q_label, q_base)
    _set_slide_note(slide, "\n".join(note_lines))
    _add_text(slide, PAGE_L, PAGE_T, PAGE_W, PAGE_H,
              str(page_num), font_size=8, color=C_QBASE,
              align=PP_ALIGN.RIGHT, anchor="ctr")


# ── Template loading ──────────────────────────────────────────────────────────

def _load_templates(templates_dir: str):
    """Load the 4 bundled chart-style chartSpace elements from XML files."""
    base = Path(templates_dir)
    tmpl = {}
    for role in ("bar", "col", "donut", "stacked"):
        path = base / f"{role}.xml"
        if not path.exists():
            raise RuntimeError(
                f"Chart template '{path}' missing. Run "
                f"extract_chart_templates.py to regenerate.")
        tmpl[role] = parse_xml(path.read_bytes())
    return tmpl


def _default_templates_dir() -> str:
    """Return the bundled chart_templates directory inside the installed package."""
    return str(Path(__file__).parent / "chart_templates")


# ── Public API ────────────────────────────────────────────────────────────────

def _safe_print(msg: str) -> None:
    try:
        print(msg)
    except UnicodeEncodeError:
        print(msg.encode("ascii", errors="replace").decode("ascii"))


def _select_table(tables: list[dict], table_idx: int | None) -> dict:
    """Pick which chart_data.json table to render into the appendix.

    Explicit ``table_idx`` always wins (CLI ``--table N`` override).
    Otherwise auto-select by ``sub_title`` (case-insensitive): prefer a
    table named "Appendix", fall back to "General". Raises ValueError if
    neither exists — appendix generation only ever targets a single table,
    it never silently renders every table in the file.
    """
    if table_idx is not None:
        for tbl in tables:
            if tbl.get("table_index") == table_idx:
                return tbl
        raise ValueError(f"No table with table_index={table_idx} in chart_data.json")

    by_sub_title = {str(t.get("sub_title", "")).strip().lower(): t for t in tables}
    for name in ("appendix", "general"):
        if name in by_sub_title:
            return by_sub_title[name]

    available = ", ".join(repr(t.get("sub_title", "")) for t in tables) or "(none)"
    raise ValueError(
        "No table named 'Appendix' or 'General' found in chart_data.json "
        f"(available sub_title values: {available}). Add a table with "
        "sub_title \"Appendix\" (or \"General\") to datatable.json."
    )


def _chunk_breakdowns_by_col_count(breakdowns: list, max_cols: int) -> list[list]:
    """Split a list of banner groups across as few slides as possible
    (never splitting one group's own columns across two slides), with each
    slide's total column count as EVEN as possible — rather than greedily
    filling each slide to the brim before starting the next (which can
    produce a lopsided e.g. 9-and-4 split), this computes the minimum
    slide count (`ceil(total / max_cols)`) and distributes whole groups
    across that many slides using LPT (longest-processing-time-first) bin
    balancing: largest groups placed first, each into whichever slide
    currently has the fewest columns — e.g. Area(4)/Gender(2)/Age(3)/HHI(4)
    = 13 cols → 2 slides balances to 7/6 instead of greedy's 9/4."""
    total = sum(len(bd.get("columns", [])) for bd in breakdowns)
    if total <= max_cols:
        return [breakdowns]

    n_slides = -(-total // max_cols)  # ceil(total / max_cols)
    order = sorted(range(len(breakdowns)),
                    key=lambda i: -len(breakdowns[i].get("columns", [])))
    bucket_idxs: list[list[int]] = [[] for _ in range(n_slides)]
    bucket_totals = [0] * n_slides
    for i in order:
        j = min(range(n_slides), key=lambda k: bucket_totals[k])
        bucket_idxs[j].append(i)
        bucket_totals[j] += len(breakdowns[i].get("columns", []))

    chunks = [
        [breakdowns[i] for i in sorted(idxs)]
        for idxs in bucket_idxs if idxs
    ]
    chunks.sort(key=lambda chunk: breakdowns.index(chunk[0]))
    return chunks


def generate(chart_data_path: str, output_path: str, *,
             templates_dir: str | None = None, table_idx: int | None = None,
             start_page: int = 1) -> None:
    """Generate a PowerPoint appendix from chart_data.json.

    Renders exactly one table (never "all tables merged") — see
    ``_select_table`` for the selection rule.

    Args:
        chart_data_path: Path to chart_data.json produced by the pipeline.
        output_path:     Where to write the .pptx file.
        templates_dir:   Override chart template XML directory (default: bundled).
        table_idx:       If set, export this table index (0-based) instead of
                          auto-selecting by sub_title.
        start_page:      Starting page number shown in slides (default: 1).
    """
    data = json.loads(Path(chart_data_path).read_text(encoding="utf-8"))
    tmpl = _load_templates(templates_dir or _default_templates_dir())
    tbl = _select_table(data.get("tables", []), table_idx)

    prs = Presentation()
    prs.slide_width  = SW
    prs.slide_height = SH
    blank_layout = prs.slide_layouts[6]

    page = start_page
    for q in tbl.get("questions", []):
        if q.get("total", {}).get("base", 0) == 0:
            continue
        _safe_print(f"  slide {page:>3}: {q.get('question', ''):<12}  {q.get('label', '')}")
        chart_type = q.get("chart_type", "bar_vertical")
        breakdowns = q.get("breakdowns", [])

        if chart_type == "donut_stacked":
            total_cols = sum(len(bd.get("columns", [])) for bd in breakdowns)
            if total_cols <= STACK_MAX_COLS_PER_SLIDE:
                _build_slide(prs, blank_layout, q, page, tmpl)
                page += 1
            else:
                # Split the stack's columns across continuation slides
                # (whole banner groups per slide, see
                # _chunk_breakdowns_by_col_count) — the Total donut is
                # re-rendered identically on every one of them.
                chunks = _chunk_breakdowns_by_col_count(breakdowns, STACK_MAX_COLS_PER_SLIDE)
                for ci, chunk in enumerate(chunks):
                    q_slide = {**q, "breakdowns": chunk}
                    _build_slide(prs, blank_layout, q_slide, page, tmpl,
                                 title_suffix=f"({ci + 1})")
                    page += 1
        elif len(breakdowns) <= MAX_GROUPS_PER_SLIDE:
            _build_slide(prs, blank_layout, q, page, tmpl)
            page += 1
        else:
            # Split into continuation slides, max MAX_GROUPS_PER_SLIDE groups each
            chunks = [breakdowns[i:i + MAX_GROUPS_PER_SLIDE]
                      for i in range(0, len(breakdowns), MAX_GROUPS_PER_SLIDE)]
            for ci, chunk in enumerate(chunks):
                q_slide = {**q, "breakdowns": chunk}
                _build_slide(prs, blank_layout, q_slide, page, tmpl,
                             title_suffix=f"({ci + 1})")
                page += 1

    prs.save(output_path)
    _safe_print(f"\nSaved -> {output_path}  ({page - start_page} slides)")


# ── CLI entry point ───────────────────────────────────────────────────────────

def main(argv=None) -> None:
    ap = argparse.ArgumentParser(
        description="Generate editable PowerPoint slides from chart_data.json",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("chart_data", help="Path to chart_data.json")
    ap.add_argument("output", help="Output .pptx path")
    ap.add_argument("--templates", default=None,
                    help="Chart-template dir (default: bundled package templates)")
    ap.add_argument("--table", type=int, default=None,
                    help="Export this table index (default: auto-select by "
                         "sub_title — 'Appendix', else 'General')")
    ap.add_argument("--start-page", type=int, default=1, dest="start_page",
                    help="Starting page number (default: 1)")
    args = ap.parse_args(argv)
    generate(args.chart_data, args.output,
             templates_dir=args.templates,
             table_idx=args.table, start_page=args.start_page)


if __name__ == "__main__":
    main()
