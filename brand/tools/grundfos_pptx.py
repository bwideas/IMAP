"""Helpers for building Grundfos-branded PowerPoint reports with python-pptx.

Everything follows the Grundfos 2022 template (fonts, colours, layouts) and the
report style described in brand/README.md. Logos and pictures are always added as
ordinary slide objects so they can be moved, resized, replaced or deleted.

Quick start:

    from grundfos_pptx import *
    prs = load_template()                                   # empty deck on the Grundfos template
    cover = new_cover(prs, "Report title", ["Subtitle", "NAS Marketing  |  Month 2026"])
    s = new_content_slide(prs, "Slide title")
    table(s, 0.58, 1.4, 12.17, [3, 3, 3, 3.17], [["Header", ...], ["row", ...]])
    prs.save("my_report.pptx")

Units: all x / y / width / height arguments are inches on a 13.333 x 7.5 in slide.
Safe content area on a content slide: x 0.58 - 12.75, y 1.3 - 6.6.
"""
from __future__ import annotations

import copy
import subprocess
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

__all__ = [
    "TEMPLATE", "DARK_BLUE", "BRIGHT_BLUE", "LIGHT_BLUE", "NAVY", "PALE_BLUE", "STEEL_BLUE",
    "LINE_GREY", "TEXT", "MUTED", "WHITE", "BODY", "BOLD", "TITLE_FONT",
    "Inches", "Pt", "PP_ALIGN", "MSO_ANCHOR",
    "load_template", "new_cover", "new_content_slide", "textbox", "label", "panel", "table",
    "bullets", "bar_chart", "add_picture", "export_png", "fmt",
]

TEMPLATE = Path(__file__).resolve().parents[1] / "templates" / "Grundfos_2022_Report_Template.pptx"

# Grundfos 2022 theme colours
DARK_BLUE = RGBColor(0x11, 0x49, 0x7B)    # accent 1 - table headers, bars, labels
BRIGHT_BLUE = RGBColor(0x12, 0x6A, 0xF3)  # accent 2 - slide titles (comes from the layout)
LIGHT_BLUE = RGBColor(0x47, 0xA6, 0xFF)   # accent 3 - use sparingly
NAVY = RGBColor(0x09, 0x28, 0x44)         # accent 4 - dark backgrounds
PALE_BLUE = RGBColor(0xE9, 0xF1, 0xF7)    # panels and total rows
STEEL_BLUE = RGBColor(0x74, 0x97, 0xB2)   # secondary text on dark slides
LINE_GREY = RGBColor(0xCC, 0xCF, 0xD7)    # table rules
TEXT = RGBColor(0x00, 0x00, 0x00)
MUTED = RGBColor(0x59, 0x59, 0x59)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# Grundfos fonts (must be installed on the machine that opens the file)
BODY = "Grundfos TheSans SemiLight"
BOLD = "Grundfos TheSans SemiBold"
TITLE_FONT = "Grundfos TheSans Extd Black"

_LAYOUT_COVER = "Title Slide"
_LAYOUT_CONTENT = "Title Only"
_MASTER_MARK = "Graphic 6"              # small "X" mark in the slide master
_COVER_PHOTO = "Picture Placeholder 6"  # full-bleed water photo in the Title Slide layout
_COVER_LOGO = "Graphic 9"               # logo + "Possibility in every drop" in the Title Slide layout


def fmt(value) -> str:
    """Thousands separator: 12345 -> '12,345'."""
    return f"{value:,}"


# --------------------------------------------------------------------------- template
def load_template(path: str | Path | None = None) -> Presentation:
    """Open the Grundfos template and remove its slides, keeping master, layouts and theme."""
    prs = Presentation(str(path or TEMPLATE))
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        prs.part.drop_rel(sldId.rId)
        sldIdLst.remove(sldId)
    return prs


def _layout(prs, name):
    return next(l for l in prs.slide_layouts if l.name == name)


def _shape(owner, name):
    return next(s for s in owner.shapes if s.name == name)


def _copy_picture(src_shape, slide, name, to_back=False):
    """Copy a master/layout picture onto the slide as a normal picture (keeps PNG + SVG)."""
    src_part = src_shape.part
    el = copy.deepcopy(src_shape._element)
    for node in el.iter():
        for attr in (qn("r:embed"), qn("r:link")):
            rId = node.get(attr)
            if rId:
                rel = src_part.rels[rId]
                node.set(attr, slide.part.relate_to(rel.target_part, rel.reltype))
    cNvPr = el.find(qn("p:nvPicPr") + "/" + qn("p:cNvPr"))
    cNvPr.set("id", str(slide.shapes._next_shape_id))
    cNvPr.set("name", name)
    spTree = slide.shapes._spTree
    if to_back:
        spTree.insert(2, el)
    else:
        spTree.append(el)
    return el


def new_cover(prs, title, subtitle_lines=(), photo: str | Path | None = None):
    """Cover slide: dark water photo, white title, subtitle lines, logo with endline.

    Pass `photo` to use another full-bleed picture instead of the template photo.
    """
    layout = _layout(prs, _LAYOUT_COVER)
    slide = prs.slides.add_slide(layout)
    slide._element.set("showMasterSp", "0")          # hide the locked layout graphics
    if photo:
        pic = slide.shapes.add_picture(str(photo), 0, 0, prs.slide_width, prs.slide_height)
        pic.name = "Cover photo"
        slide.shapes._spTree.remove(pic._element)
        slide.shapes._spTree.insert(2, pic._element)
    else:
        _copy_picture(_shape(layout, _COVER_PHOTO), slide, "Cover photo", to_back=True)
    _copy_picture(_shape(layout, _COVER_LOGO), slide, "Grundfos logo")
    slide.shapes.title.text = title
    sub = slide.placeholders[1]
    sub.left, sub.top, sub.width, sub.height = Inches(0.58), Inches(3.62), Inches(9.3), Inches(1.0)
    lines = list(subtitle_lines)
    sub.text_frame.text = lines[0] if lines else ""
    for line in lines[1:]:
        sub.text_frame.add_paragraph().text = line
    return slide


def new_content_slide(prs, title):
    """White slide with blue title (from the layout) and a movable Grundfos mark top-right."""
    layout = _layout(prs, _LAYOUT_CONTENT)
    slide = prs.slides.add_slide(layout)
    slide._element.set("showMasterSp", "0")
    _copy_picture(_shape(prs.slide_master, _MASTER_MARK), slide, "Grundfos mark")
    slide.shapes.title.text = title
    return slide


# --------------------------------------------------------------------------- text
def _para_runs(para, runs, size, color=TEXT, align=None, space_after=None, line_spacing=None):
    if align is not None:
        para.alignment = align
    if space_after is not None:
        para.space_after = Pt(space_after)
    if line_spacing is not None:
        para.line_spacing = line_spacing
    for item in runs:
        txt, opts = (item, {}) if isinstance(item, str) else item
        r = para.add_run()
        r.text = txt
        r.font.name = BOLD if opts.get("bold") else BODY
        r.font.size = Pt(opts.get("size", size))
        r.font.color.rgb = opts.get("color", color)
        r.font.bold = False   # weight comes from the font family, never synthetic bold


def textbox(slide, x, y, w, h, paragraphs, size=11, color=TEXT, align=None,
            anchor=MSO_ANCHOR.TOP, space_after=4, line_spacing=1.05):
    """Add a text box.

    `paragraphs` is a list; each item is a string or a list of runs, where a run is a
    string or ("text", {"bold": True, "color": DARK_BLUE, "size": 12}).
    """
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    for i, spec in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        runs = [spec] if isinstance(spec, str) else spec
        _para_runs(p, runs, size, color, align, space_after, line_spacing)
    return tb


def label(slide, x, y, w, text, size=12, color=DARK_BLUE):
    """Small dark-blue heading, e.g. above a chart or inside a panel."""
    return textbox(slide, x, y, w, 0.3, [[(text, {"bold": True, "color": color})]], size=size)


def bullets(slide, x, y, w, h, items, size=11.5, space_after=7):
    """Plain bullet list. Each item is a string or a list of runs (see textbox)."""
    paras = []
    for it in items:
        runs = [("\u2022  ", {"color": DARK_BLUE})]
        runs += [it] if isinstance(it, str) else it
        paras.append(runs)
    return textbox(slide, x, y, w, h, paras, size=size, space_after=space_after, line_spacing=1.08)


# --------------------------------------------------------------------------- shapes
def panel(slide, x, y, w, h, fill=PALE_BLUE):
    """Flat pale-blue rectangle used behind 'what stood out' / 'next steps' text."""
    shp = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_picture(slide, path, x, y, w=None, h=None, name="Picture"):
    kw = {}
    if w is not None:
        kw["width"] = Inches(w)
    if h is not None:
        kw["height"] = Inches(h)
    pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y), **kw)
    pic.name = name
    return pic


# --------------------------------------------------------------------------- tables
def _cell_borders(cell, bottom=LINE_GREY, width_pt=0.75):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for old in tcPr.findall(qn(tag)):
            tcPr.remove(old)
    nodes = []
    for tag in ("a:lnL", "a:lnR", "a:lnT"):
        ln = etree.SubElement(etree.Element("dummy"), qn(tag), w="0")
        etree.SubElement(ln, qn("a:noFill"))
        nodes.append(ln)
    lnB = etree.SubElement(etree.Element("dummy"), qn("a:lnB"), w=str(int(width_pt * 12700)),
                           cap="flat", cmpd="sng", algn="ctr")
    sf = etree.SubElement(lnB, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr"), val=str(bottom))
    etree.SubElement(lnB, qn("a:prstDash"), val="solid")
    nodes.append(lnB)
    for i, node in enumerate(nodes):       # line elements must precede the fill in a:tcPr
        tcPr.insert(i, node)


def table(slide, x, y, w, col_w, rows, size=10.5, row_h=0.42, header_h=None, total_row=True,
          align=None, bold_first_col=False):
    """Grundfos-style table: dark-blue header, grey horizontal rules, pale-blue total row.

    rows     - list of rows; first row is the header. A cell is a string or a list of runs.
    col_w    - column widths in inches (should sum to w).
    align    - {col_index: PP_ALIGN.RIGHT, ...}
    """
    nrows, ncols = len(rows), len(rows[0])
    gf = slide.shapes.add_table(nrows, ncols, Inches(x), Inches(y), Inches(w), Inches(row_h * nrows))
    tbl = gf.table
    tbl.first_row, tbl.horz_banding, tbl.vert_banding = True, False, False
    for i, cw in enumerate(col_w):
        tbl.columns[i].width = Inches(cw)
    for r in range(nrows):
        tbl.rows[r].height = Inches(header_h if (r == 0 and header_h) else row_h)
        is_header, is_total = r == 0, total_row and r == nrows - 1
        for c in range(ncols):
            cell = tbl.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            _cell_borders(cell, bottom=DARK_BLUE if is_header else LINE_GREY)
            cell.fill.solid()
            cell.fill.fore_color.rgb = DARK_BLUE if is_header else (PALE_BLUE if is_total else WHITE)
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            if align and c in align:
                p.alignment = align[c]
            val = rows[r][c]
            runs = val if isinstance(val, list) else [val]
            emphasise = is_header or is_total or (bold_first_col and c == 0)
            runs = [(t, {"bold": True, **o}) if emphasise else (t, o)
                    for t, o in ((it, {}) if isinstance(it, str) else it for it in runs)]
            _para_runs(p, runs, size, WHITE if is_header else TEXT)
    return tbl


# --------------------------------------------------------------------------- charts
def bar_chart(slide, x, y, w, h, categories, values, number_format="0%", max_scale=None,
              size=10.5, largest_on_top=True):
    """Single-colour horizontal bar chart with value labels; no gridlines or legend."""
    cats, vals = list(categories), list(values)
    if largest_on_top:          # bar charts draw bottom-up, so reverse the order
        cats, vals = cats[::-1], vals[::-1]
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("Series", vals)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), cd)
    chart = gf.chart
    chart.has_title = False
    chart.has_legend = False
    chart.font.name = BODY
    chart.font.size = Pt(size)
    chart.font.color.rgb = TEXT
    plot = chart.plots[0]
    plot.gap_width = 60
    plot.has_data_labels = True
    plot.data_labels.number_format = number_format
    plot.data_labels.number_format_is_linked = False
    plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    plot.data_labels.font.size = Pt(size)
    plot.data_labels.font.name = BOLD
    ser = plot.series[0]
    ser.format.fill.solid()
    ser.format.fill.fore_color.rgb = DARK_BLUE
    va = chart.value_axis
    va.visible = False
    va.has_major_gridlines = False
    va.minimum_scale = 0
    if max_scale is not None:
        va.maximum_scale = max_scale
    ca = chart.category_axis
    ca.format.line.color.rgb = LINE_GREY
    ca.has_major_gridlines = False
    ca.tick_labels.font.size = Pt(size)
    ca.tick_labels.font.name = BODY
    return chart


# --------------------------------------------------------------------------- preview
def export_png(pptx_path, out_dir, width=1600, height=900):
    """Export every slide to PNG with PowerPoint (Windows only). Returns the list of files."""
    pptx_path, out_dir = Path(pptx_path).resolve(), Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob("slide*.png"):
        old.unlink()
    script = f"""
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open('{pptx_path}', $true, $false, $false)
$i = 1
foreach ($s in $pres.Slides) {{
    $s.Export(('{out_dir}\\slide' + $i.ToString('00') + '.png'), 'PNG', {width}, {height}); $i++
}}
$pres.Close(); $app.Quit()
"""
    subprocess.run(["powershell", "-NoProfile", "-Command", script], check=True)
    return sorted(out_dir.glob("slide*.png"))
