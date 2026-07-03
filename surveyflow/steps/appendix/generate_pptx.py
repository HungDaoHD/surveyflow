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
    from pptx.enum.text import PP_ALIGN
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

# Above this label length, PowerPoint tends to render a chart's legend as a
# single vertical column (one entry per row) instead of flowing multiple
# entries per row. Confirmed via real PowerPoint rendering.
_LEGEND_SINGLE_COL_LEN = 35


def _likely_single_column_legend(labels) -> bool:
    """Heuristic for whether a legend will render as a single vertical
    column. Matters specifically for STACKED charts: confirmed via real
    PowerPoint rendering that a single-column legend displays in the
    REVERSE of series-add order (matching the visual top-to-bottom
    stacking), while a multi-column (multiple entries per row) legend does
    not reverse. There's no way to precisely predict PowerPoint's own text
    layout from Python, so this uses a practical length threshold rather
    than exact measurement."""
    return any(len(lbl) > _LEGEND_SINGLE_COL_LEN for lbl in labels)


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

def _style_cat_axis(chartspace, *, font_pt: int = 8) -> None:
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


def _style_legend(chartspace, n_categories: int, *, base_pt: int = 9, min_pt: int = 6) -> None:
    """Shrink legend text as the category count grows and force word-wrap, so
    every visible entry stays fully readable.

    PowerPoint's legend area has a fixed height/width — long entry text that
    doesn't fit gets clipped instead of wrapping, and with many entries at the
    template's default 9pt some can be dropped entirely rather than shrunk.
    Enabling wrap + scaling the font down keeps every entry visible."""
    font_pt = base_pt if n_categories <= 6 else max(min_pt, base_pt - (n_categories - 6))
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
                 palette: list | None = None, style_cat_labels: bool = False,
                 cat_font_pt: int = 8, legend_n: int | None = None,
                 hide_zero_labels: bool = False, series_colors: list | None = None):
    """Add a chart, replace its XML with a deep copy of the template chartSpace,
    then inject `chart_data` via replace_data() so template styling is kept."""
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
    if style_cat_labels:
        _style_cat_axis(cp._element, font_pt=cat_font_pt)
    if legend_n is not None:
        _ensure_legend(cp._element)
        _style_legend(cp._element, legend_n)
    if per_point:
        _match_point_dlbl_count(cp._element)
    if hide_zero_labels:
        _hide_zero_value_labels(cp._element)
    _apply_palette(cp._element, palette or CHART_PALETTE, per_point=per_point,
                   series_colors=series_colors)
    return cp.chart


def _v100(percents: dict, codes: list) -> list:
    """Fractions (0-1) to percentages (0-100) for the given choice codes."""
    return [round(percents.get(c, 0.0) * 100.0, 4) for c in codes]


def _col_label_with_base(col: dict) -> str:
    """Breakdown column label + its own base, e.g. 'Male (N=120)'."""
    return f"{col.get('label', '')} (N={col.get('base', 0)})"


def _sort_scale_desc(q: dict, choices: list) -> list:
    """SA-Ordinal questions (`scale_class == "Ordinal"`, Claude-classified in
    metadata.json — see CLAUDE.md Step 3c): sort choices highest scale point
    to lowest. SA-Nominal questions are returned unchanged.

    Which code is "highest" isn't always the max code value — e.g. a
    frequency scale coded 1=Almost every day … 7=Less than once a month has
    code 1 at the high-intensity end. `scale_high_code` (also
    Claude-classified, see CLAUDE.md Step 3c) names the code at the high end
    explicitly; sort direction is derived from whether that's the min or max
    code. Defaults to "max code = high end" (descending by code) when
    `scale_high_code` isn't set."""
    if q.get("scale_class") != "Ordinal":
        return choices
    try:
        codes = [int(c["code"]) for c in choices]
    except (ValueError, TypeError):
        return list(reversed(choices))
    high_code = q.get("scale_high_code")
    reverse = True
    if high_code is not None:
        try:
            reverse = int(high_code) != min(codes)
        except (ValueError, TypeError):
            pass
    return sorted(choices, key=lambda c: int(c["code"]), reverse=reverse)


def _order_sa_choices(q: dict, choices: list) -> list:
    """Order SA (donut_stacked) choices for both the donut and its paired
    stacked breakdown chart, so the two always agree:

    - SA-Ordinal: sort by scale order descending (see _sort_scale_desc).
    - SA-Nominal: sort by Total percent descending."""
    if q.get("scale_class") == "Ordinal":
        return _sort_scale_desc(q, choices)
    total_pct = q.get("total", {}).get("percents", {})
    return sorted(choices, key=lambda c: total_pct.get(c["code"], 0), reverse=True)


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
    _clone_chart(slide, tmpl["col"], XL_CHART_TYPE.COLUMN_CLUSTERED, l, t, w, h, cd,
                 style_cat_labels=True)


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
                 style_cat_labels=True,
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
                     l, t, w, h, cd, style_cat_labels=True, cat_font_pt=7)
    else:
        _clone_chart(slide, tmpl["stacked"], XL_CHART_TYPE.COLUMN_STACKED_100,
                     l, t, w, h, cd, style_cat_labels=True, cat_font_pt=7)


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


def _build_stacked(slide, q, breakdowns, tmpl, l, t, w, h) -> None:
    """100%-stacked columns — series per choice, using the same DONUT_PALETTE
    colors per choice so the two charts agree. Shows every choice (incl.
    those that are 0% everywhere) so the legend always matches the donut's
    full list.

    Legend order (SA-Nominal only): confirmed via real PowerPoint rendering
    that a multi-column legend (short labels) displays in literal series-add
    order, while a single-column legend (long labels, wraps one-per-row)
    displays REVERSED — so for SA-Nominal questions, series are added
    choices-reversed only when _likely_single_column_legend() predicts the
    latter, keeping the rendered legend matching the donut's order either
    way. SA-Ordinal questions skip this compensation and always use plain
    scale order, since their order is meaningful on its own (not just a
    percent ranking) and shouldn't be perturbed by a label-length heuristic."""
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
    is_nominal = q.get("scale_class") != "Ordinal"
    insertion_order = (
        list(reversed(choices))
        if is_nominal and _likely_single_column_legend(shortened.values())
        else choices
    )
    cd = CategoryChartData()
    cd.categories = xlabels
    series_colors = []
    for choice in insertion_order:
        cd.add_series(
            shortened[choice["code"]],
            [round(col.get("percents", {}).get(choice["code"], 0.0) * 100.0, 4)
             for col in all_cols],
        )
        series_colors.append(color_map[choice["code"]])
    _clone_chart(slide, tmpl["stacked"], XL_CHART_TYPE.COLUMN_STACKED_100, l, t, w, h, cd,
                 style_cat_labels=True, cat_font_pt=7, series_colors=series_colors,
                 hide_zero_labels=True)


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


def generate(chart_data_path: str, output_path: str, *,
             templates_dir: str | None = None, table_idx: int | None = None,
             start_page: int = 1) -> None:
    """Generate a PowerPoint appendix from chart_data.json.

    Args:
        chart_data_path: Path to chart_data.json produced by the pipeline.
        output_path:     Where to write the .pptx file.
        templates_dir:   Override chart template XML directory (default: bundled).
        table_idx:       If set, only export this table index (0-based).
        start_page:      Starting page number shown in slides (default: 1).
    """
    data = json.loads(Path(chart_data_path).read_text(encoding="utf-8"))
    tmpl = _load_templates(templates_dir or _default_templates_dir())

    prs = Presentation()
    prs.slide_width  = SW
    prs.slide_height = SH
    blank_layout = prs.slide_layouts[6]

    page = start_page
    for tbl in data.get("tables", []):
        if table_idx is not None and tbl.get("table_index") != table_idx:
            continue
        for q in tbl.get("questions", []):
            if q.get("total", {}).get("base", 0) == 0:
                continue
            _safe_print(f"  slide {page:>3}: {q.get('question', ''):<12}  {q.get('label', '')}")
            chart_type = q.get("chart_type", "bar_vertical")
            breakdowns = q.get("breakdowns", [])

            # donut_stacked: all breakdowns go into one stacked chart — no split needed
            if chart_type == "donut_stacked" or len(breakdowns) <= MAX_GROUPS_PER_SLIDE:
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
                    help="Export only this table index (default: all)")
    ap.add_argument("--start-page", type=int, default=1, dest="start_page",
                    help="Starting page number (default: 1)")
    args = ap.parse_args(argv)
    generate(args.chart_data, args.output,
             templates_dir=args.templates,
             table_idx=args.table, start_page=args.start_page)


if __name__ == "__main__":
    main()
