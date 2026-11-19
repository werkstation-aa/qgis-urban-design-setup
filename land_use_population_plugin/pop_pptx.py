"""PowerPoint output: the compact population assumption table on the 'Population Assumption'
template (A3, layout Main_blank): black header, merged 'Assumption for' and 'Land Use' cells
coloured per group, white row lines. Long tables continue on further slides.

Needs python-pptx (shipped in the plugin's vendor/ folder); lxml comes with QGIS.
"""

import copy
import os

from . import pop_core as core

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "population_template.pptx")
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

# geometry of the example slide (inches, A3 16.53 x 11.69); scaled to the template's width
TABLE_X, TABLE_Y, TABLE_W = 0.59, 2.38, 15.35
COL_W = [1.99, 1.99, 2.88, 1.85, 1.23, 5.40]
BOTTOM = 10.45  # keep clear of the footnote placeholder
ROW_H, HEAD_H = 0.30, 0.30


def _q(tag):
    return "{%s}%s" % (A, tag)


def _sub(parent, tag, **attrs):
    from lxml import etree
    el = etree.SubElement(parent, _q(tag))
    for k, v in attrs.items():
        el.set(k, str(v))
    return el


def _style_cell(cell, fill=None, line_top=True, line_bottom=True, theme_fill=None):
    """Rewrite tcPr: margins, centred, white 1 pt top/bottom lines, no side lines, fill."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for ch in list(tcPr):
        tcPr.remove(ch)
    for k, v in (("marL", 45720), ("marR", 45720), ("marT", 18288), ("marB", 18288), ("anchor", "ctr")):
        tcPr.set(k, str(v))
    for side, on in (("lnL", False), ("lnR", False), ("lnT", line_top), ("lnB", line_bottom)):
        ln = _sub(tcPr, side)
        if on:
            ln.set("w", "12700")
            sf = _sub(ln, "solidFill")
            _sub(sf, "schemeClr", val="bg1")
        else:
            _sub(ln, "noFill")
    if theme_fill:
        sf = _sub(tcPr, "solidFill")
        _sub(sf, "schemeClr", val=theme_fill)
    elif fill:
        sf = _sub(tcPr, "solidFill")
        _sub(sf, "srgbClr", val=fill)


def _text(cell, runs, align="l"):
    """runs: [(text, dict(size, bold, color, italic))]"""
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Pt
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = {"l": PP_ALIGN.LEFT, "r": PP_ALIGN.RIGHT, "c": PP_ALIGN.CENTER}[align]
    for text, st in runs:
        r = p.add_run()
        r.text = text
        f = r.font
        f.size = Pt(st.get("size", 11))
        f.bold = st.get("bold", False)
        f.italic = st.get("italic", False)
        f.color.rgb = RGBColor.from_string(st.get("color", "000000"))


def _clear_slides(prs):
    lst = prs.slides._sldIdLst
    for sld in list(lst):
        prs.part.drop_rel(sld.rId)
        lst.remove(sld)


def _layout(prs, src_layout_name):
    for lay in prs.slide_layouts:
        if lay.name == src_layout_name:
            return lay
    return prs.slide_layouts[len(prs.slide_layouts) - 1]


def _fill_placeholders(slide, title, footnote, scale):
    from pptx.util import Inches, Pt
    def widen(ph):
        """Placeholder as wide as the table. Position is inherited from the layout until the slide
        gets its own xfrm, so all four values are written together."""
        x, y, h = ph.left, ph.top, ph.height
        ph.left, ph.top, ph.width, ph.height = x, y, Inches(TABLE_W * scale), h

    done_title = done_foot = False
    for ph in slide.placeholders:
        t = ph.placeholder_format.type
        if not done_title and "TITLE" in str(t):
            ph.text_frame.text = title
            widen(ph)
            done_title = True
        elif not done_foot and "BODY" in str(t):
            ph.text_frame.text = footnote
            widen(ph)
            done_foot = True
    if not done_title:
        tb = slide.shapes.add_textbox(Inches(0.59 * scale), Inches(1.06 * scale), Inches(10 * scale), Inches(0.6))
        tb.text_frame.text = title
        tb.text_frame.paragraphs[0].runs[0].font.size = Pt(28)
        tb.text_frame.paragraphs[0].runs[0].font.bold = True
    if not done_foot and footnote:
        tb = slide.shapes.add_textbox(Inches(0.59 * scale), Inches(10.71 * scale), Inches(15 * scale), Inches(0.3))
        tb.text_frame.text = footnote
        tb.text_frame.paragraphs[0].runs[0].font.size = Pt(10)


def write_pptx(path, rows, title="Population Assumption", footnote="", template=None, max_rows=None):
    """rows from pop_core.compact_rows(). Returns the number of slides written."""
    from pptx import Presentation
    from pptx.util import Emu, Inches

    tpl = template if template and os.path.exists(template) else TEMPLATE
    prs = Presentation(tpl)
    layout_name = prs.slides[0].slide_layout.name if len(prs.slides) else "Main_blank"
    _clear_slides(prs)
    layout = _layout(prs, layout_name)
    scale = Emu(prs.slide_width).inches / 16.535
    y0, bottom = TABLE_Y * scale, BOTTOM * Emu(prs.slide_height).inches / 11.693
    per_slide = max_rows or max(4, int((bottom - y0 - HEAD_H) / ROW_H))
    chunks = [rows[i:i + per_slide] for i in range(0, len(rows), per_slide)] or [[]]

    for n, chunk in enumerate(chunks):
        slide = prs.slides.add_slide(layout)
        _fill_placeholders(slide, title + (" (continued)" if n else ""), footnote, scale)
        nrows = len(chunk) + 1
        gf = slide.shapes.add_table(nrows, 6, Inches(TABLE_X * scale), Inches(y0), Inches(TABLE_W * scale),
                                    Inches(HEAD_H + ROW_H * len(chunk)))
        gf.name = "Population Assumption Table"
        tbl = gf.table
        tblPr = tbl._tbl.tblPr
        for attr in list(tblPr.attrib):  # no banding / first-row style: every cell is styled explicitly
            del tblPr.attrib[attr]
        for ch in list(tblPr):
            tblPr.remove(ch)
        for i, w in enumerate(COL_W):
            tbl.columns[i].width = Inches(w * scale)
        tbl.rows[0].height = Inches(HEAD_H)
        for i in range(1, nrows):
            tbl.rows[i].height = Inches(ROW_H)

        for j, lab in enumerate(core.COMPACT_COLS):
            c = tbl.cell(0, j)
            _style_cell(c, fill="000000", line_top=False, line_bottom=j >= 2)
            _text(c, [(lab, dict(bold=True, color="FFFFFF"))])

        for i, r in enumerate(chunk, 1):
            vals = [r["coef"], r["units"]]
            c = tbl.cell(i, 2)
            _style_cell(c, fill="FFFFFF")
            _text(c, [(r["program"], {})])
            c = tbl.cell(i, 3)
            _style_cell(c)
            _text(c, [(vals[0], {})])
            c = tbl.cell(i, 4)
            _style_cell(c)
            _text(c, [(vals[1], {})])
            c = tbl.cell(i, 5)
            _style_cell(c)
            _text(c, [(r["source"], {}), (f"  · {r['status']}", dict(size=9, italic=True, color="808080"))])

        for key, col in (("assumption_for", 0), ("land_use", 1)):
            for a, b in core.spans(chunk, key):
                r = chunk[a]
                hexc = r["for_color"] if key == "assumption_for" else r["color"]
                theme = "accent4" if key == "assumption_for" and r["assumption_for"] != "Resident" else None
                for i in range(a, b + 1):
                    _style_cell(tbl.cell(i + 1, col), fill=hexc, theme_fill=theme, line_top=False, line_bottom=False)
                if b > a:
                    tbl.cell(a + 1, col).merge(tbl.cell(b + 1, col))
                dark = (not theme) and core.is_dark(hexc)
                _text(tbl.cell(a + 1, col), [(r[key], dict(bold=True, color="FFFFFF" if dark else "000000"))])
    prs.save(path)
    return len(chunks)
