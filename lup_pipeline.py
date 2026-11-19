"""
Land Use Plan automation: Shapefile -> Excel database -> PowerPoint tables/charts.

    py lup_pipeline.py                                  # uses defaults below
    py lup_pipeline.py --shp "Shapefile\\X.shp" --site-total 5000000 --out Output

Steps
  1. Read the land-use polygons and their attributes (pyshp, no GDAL needed).
  2. Write an Excel workbook:
       Source      raw attribute table (+ geometry area check)
       Overall MP  Saleable / Non-Saleable summary, industrial breakdown, population
       Phasing     area per phase x SUBZONA, plus Green / Parcel / Road per phase
       Intensitas  parcel intensity (BCR, FAR, height, GOSR, GFA = FAR x area)
       Checks      data-quality warnings
     Summary sheets use SUMIFS formulas on Source, so pasting new data into
     Source updates them in Excel.
  3. Write a PowerPoint deck with the two table + chart slides from the sample.
"""

import argparse
import math
import os
import sys
from collections import OrderedDict, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))

# --------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------
def _first_shp(folder):
    try:
        return next(os.path.join(folder, f) for f in sorted(os.listdir(folder)) if f.lower().endswith(".shp"))
    except (OSError, StopIteration):
        return os.path.join(folder, "LandUse.shp")


DEFAULT_SHP = _first_shp(os.path.join(HERE, "Shapefile"))  # first shapefile in Shapefile\
DEFAULT_OUT = os.path.join(HERE, "Output")
DEFAULT_TEMPLATE = os.path.join(HERE, "LUP_template.pptx")

# Site area used for GRAND TOTAL. Road takes the remainder of the site total.
# Set to None to use the plain sum of the shapefile instead.
SITE_TOTAL_SQM = 5_000_000

# Number style for the slides: "id" -> 3.151.645 / 315,16   "en" -> 3,151,645 / 315.16
NUMBER_STYLE = "id"

# Attribute field names in the shapefile
F_TYPE, F_SUBZ, F_INDT = "Type", "SUBZONA", "IND_TYPE"
F_AREA, F_PHASE, F_PARCEL = "Area_sqm", "Phase", "Parcel_ID"
F_POP, F_POPTYPE = "Pop_Total", "Pop_Type"
ROAD = "Road"
SALEABLE, NON_SALEABLE = "Saleable", "Non-Saleable"

# Table order (unknown categories are appended at the end of their group)
SALEABLE_ORDER = ["Industrial", "Hotel", "Commercial", "MICE", "Residential",
                  "Dormitory", "Office", "Port"]
NON_SALEABLE_ORDER = ["Public Amenity", "Golf", "Green", "Road", "Utility"]
# Order of slices in the pie charts (clockwise from 12 o'clock)
CHART_ORDER = ["Golf", "Green", "Public Amenity", "Road", "Utility",
               "Commercial", "Dormitory", "Hotel", "Industrial", "MICE",
               "Residential", "Office", "Port"]

COLORS = {  # land-use colours from the sample deck
    "Industrial": "ADDDF7", "Hotel": "CD0000", "Commercial": "FF6D6D",
    "MICE": "833232", "Residential": "FFEF82", "Dormitory": "FFDF00",
    "Office": "BA3B5D", "Port": "544CC3", "Public Amenity": "9963CC",
    "Golf": "346E2B", "Green": "58BA48", "Road": "FFFFFF", "Utility": "999999",
    # industrial types
    "Light Industry": "43AFD0", "Medium Industry": "A63A23",
    "Heavy Industry": "DAD41C", "Agro Processing": "8B5729",
    "Food Processing": "CA8B5A", "Logistics": "DFC664",
}
FALLBACK_COLORS = ["4E79A7", "F28E2B", "76B7B2", "EDC948", "B07AA1",
                   "FF9DA7", "9C755F", "BAB0AC", "59A14F", "E15759"]
GROUP_COLORS = {SALEABLE: "B89128", NON_SALEABLE: "92D050"}
SALEABLE_HDR, NON_SALEABLE_HDR, SUBTOTAL_GREY = "767171", "92D050", "BFBFBF"

SHORT = {  # chart labels
    "Industrial": "Industry", "Residential": "Resi", "Public Amenity": "Amenity",
    "Light Industry": "Light", "Medium Industry": "Medium", "Heavy Industry": "Heavy",
    "Agro Processing": "Agro", "Food Processing": "Food",
}

_fallback = iter(FALLBACK_COLORS * 10)
BASE_COLORS = dict(COLORS)


def set_palette(mapping=None):
    """Use `mapping` {category: HEX} on top of the project palette (from lup_colors)."""
    global _fallback
    COLORS.clear()
    COLORS.update(BASE_COLORS)
    COLORS.update(mapping or {})
    _fallback = iter(FALLBACK_COLORS * 10)


def color_of(name):
    if name not in COLORS:
        COLORS[name] = next(_fallback)
    return COLORS[name]


def is_dark(hex_):
    r, g, b = (int(hex_[i:i + 2], 16) for i in (0, 2, 4))
    return (0.299 * r + 0.587 * g + 0.114 * b) < 150


# --------------------------------------------------------------------------
# 1. READ SHAPEFILE
# --------------------------------------------------------------------------
def polygon_area(shape):
    """Planar area of a pyshp polygon (outer rings clockwise, holes counter-clockwise)."""
    pts, parts = shape.points, list(shape.parts) + [len(shape.points)]
    total = 0.0
    for a, b in zip(parts[:-1], parts[1:]):
        ring = pts[a:b]
        s = sum(x1 * y2 - x2 * y1 for (x1, y1), (x2, y2) in zip(ring, ring[1:] + ring[:1]))
        total += -s / 2.0  # clockwise outer ring -> positive
    return abs(total)


def read_shapefile(path):
    import shapefile  # pyshp; only needed when not running inside QGIS
    for enc in ("utf-8", "latin-1"):
        try:
            sf = shapefile.Reader(path, encoding=enc)
            fields = [f[0] for f in sf.fields[1:]]
            rows = []
            for sr in sf.iterShapeRecords():
                rec = dict(zip(fields, sr.record))
                rec["_geom_area"] = polygon_area(sr.shape)
                rows.append(rec)
            return fields, rows
        except UnicodeDecodeError:
            continue
    raise RuntimeError("Could not decode " + path)


def clean(v):
    if isinstance(v, str):
        v = v.strip()
        return v or None
    return v


def num(v):
    try:
        return float(v) if v not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


# --------------------------------------------------------------------------
# 2. AGGREGATE
# --------------------------------------------------------------------------
def ordered(keys, preferred):
    keys = [k for k in keys if k]
    return [k for k in preferred if k in keys] + sorted(k for k in keys if k not in preferred)


def summarise(rows, site_total):
    by = defaultdict(float)
    ind = defaultdict(float)
    pop = defaultdict(float)
    for r in rows:
        a = num(r.get(F_AREA))
        by[(clean(r.get(F_TYPE)), clean(r.get(F_SUBZ)))] += a
        if clean(r.get(F_SUBZ)) == "Industrial":
            ind[clean(r.get(F_INDT)) or "(unspecified)"] += a
        if clean(r.get(F_POPTYPE)):
            pop[clean(r.get(F_POPTYPE))] += num(r.get(F_POP))

    sale = OrderedDict((s, by[(SALEABLE, s)]) for s in
                       ordered({s for t, s in by if t == SALEABLE}, SALEABLE_ORDER))
    nons = OrderedDict((s, by[(NON_SALEABLE, s)]) for s in
                       ordered({s for t, s in by if t == NON_SALEABLE} | {ROAD}, NON_SALEABLE_ORDER))
    shp_total = sum(by.values())
    total = site_total if site_total else shp_total
    road_adj = total - shp_total          # rounding / plug difference goes to Road
    nons[ROAD] = nons.get(ROAD, 0.0) + road_adj

    other = {k: v for k, v in by.items() if k[0] not in (SALEABLE, NON_SALEABLE)}
    return dict(sale=sale, nons=nons, total=total, shp_total=shp_total, road_adj=road_adj,
                ind=OrderedDict(sorted(ind.items())), pop=OrderedDict(sorted(pop.items())),
                other=other)


def phasing_data(rows, S):
    """{phase: {SUBZONA: sqm}}; the Road adjustment goes to the last phase (as in Excel)."""
    acc = defaultdict(lambda: defaultdict(float))
    for r in rows:
        acc[clean(r.get(F_PHASE)) or "(No phase)"][clean(r.get(F_SUBZ))] += num(r.get(F_AREA))
    phases = sorted(acc)
    real = [p for p in phases if p != "(No phase)"]
    out = OrderedDict()
    for p in phases:
        subs = acc[p]
        if real and p == real[-1]:
            subs[ROAD] += S["road_adj"]
        out[p] = OrderedDict((s, subs[s]) for s in ordered(set(subs), CHART_ORDER))
    return out


def intensity_rows(rows):
    """Parcel rows for the Intensitas sheet, grouped Type > SUBZONA > Parcel_ID."""
    out = []
    for t in (SALEABLE, NON_SALEABLE):
        grp = [r for r in rows if clean(r.get(F_TYPE)) == t and clean(r.get(F_PARCEL))]
        subs = ordered({clean(r.get(F_SUBZ)) for r in grp},
                       SALEABLE_ORDER if t == SALEABLE else NON_SALEABLE_ORDER)
        for s in subs:
            for r in sorted((r for r in grp if clean(r.get(F_SUBZ)) == s), key=lambda q: str(q.get(F_PARCEL))):
                area, far, bcr = num(r.get(F_AREA)), r.get("FAR"), r.get("BCR (%)")
                out.append(dict(Type=t, SUBZONA=s, Parcel_ID=clean(r.get(F_PARCEL)),
                                Phase=clean(r.get(F_PHASE)), BCR=bcr or None, FAR=far or None,
                                MaxH=r.get("Max Height") or None, GOSR=r.get("GOSR (%)") or None,
                                Area=area, GFA=num(far) * area if far else None,
                                Footprint=num(bcr) / 100 * area if bcr else None,
                                Pop=num(r.get(F_POP)) or None))
    return out


# --------------------------------------------------------------------------
# 3. EXCEL
# --------------------------------------------------------------------------
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="000000")
HDR_FONT = Font(bold=True, color="FFFFFF")
FMT_SQM, FMT_HA, FMT_PCT = "#,##0", "#,##0.00", "0.0%"


def fill(hex_):
    return PatternFill("solid", fgColor=hex_)


def header(ws, row, col, labels):
    for i, h in enumerate(labels):
        c = ws.cell(row, col + i, h)
        c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def colored_label(cell, hex_, bold=False):
    cell.fill = fill(hex_)
    cell.font = Font(bold=bold, color="FFFFFF" if is_dark(hex_) else "000000")
    cell.border = BOX


PIVOT_SHEETS = ["Overall MP", "Industrial", "Phasing", "Intensitas", "Population"]
SOURCE_TABLE = "tblSource"


def _cf_value_rules(ws, ref, scheme):
    """Conditional formatting: one 'cell value = X' rule per category -> fill + readable font."""
    from openpyxl.formatting.rule import CellIsRule
    for value, hx in scheme.items():
        txt = '"' + str(value).replace('"', '""') + '"'
        ws.conditional_formatting.add(ref, CellIsRule(
            operator="equal", formula=[txt], fill=PatternFill("solid", fgColor=hx, bgColor=hx),
            font=Font(color="FFFFFF" if is_dark(hx) else "000000")))


def _base_workbook(fields, rows, S, colors):
    """Source (raw attribute table as an Excel table), Colours legend, Checks."""
    wb = Workbook()
    src = wb.active
    src.title = "Source"
    real = [f for f in fields if any(f in r for r in rows)]  # fields that exist in the layer
    src.append(real)
    for r in rows:
        src.append([clean(r.get(f)) if not isinstance(r.get(f), (int, float)) else r.get(f) for f in real])
    n = len(rows) + 1
    last = get_column_letter(len(real))
    from openpyxl.worksheet.table import Table, TableStyleInfo
    tbl = Table(displayName=SOURCE_TABLE, ref=f"A1:{last}{n}")
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    src.add_table(tbl)
    src.freeze_panes = "A2"
    for j, f in enumerate(real, start=1):
        src.column_dimensions[get_column_letter(j)].width = max(10, min(30, len(f) + 4))
    # colour coding by value (conditional formatting)
    col = {f: get_column_letter(j) for j, f in enumerate(real, start=1)}
    if colors["field"] in col:
        _cf_value_rules(src, f"{col[colors['field']]}2:{col[colors['field']]}{n}", colors["scheme"])
    if F_SUBZ in col and colors["field"] != F_SUBZ:
        _cf_value_rules(src, f"{col[F_SUBZ]}2:{col[F_SUBZ]}{n}", colors["groups"])
    if F_INDT in col:
        _cf_value_rules(src, f"{col[F_INDT]}2:{col[F_INDT]}{n}", {k: color_of(k) for k in S["ind"]})
    if F_TYPE in col:
        _cf_value_rules(src, f"{col[F_TYPE]}2:{col[F_TYPE]}{n}", GROUP_COLORS)

    # colour legend
    lg = wb.create_sheet("Colours")
    lg["A1"] = f"COLOUR CODING - field '{colors['field']}'"
    lg["A1"].font = Font(bold=True, size=13)
    header(lg, 3, 1, [colors["field"], "Colour", "HEX", "RGB", "Source", "Library match", "Features", "Area (sqm)"])
    cnt, area = defaultdict(int), defaultdict(float)
    for r in rows:
        v = r.get(colors["field"])
        if v not in (None, ""):
            cnt[str(v).strip()] += 1
            area[str(v).strip()] += num(r.get(F_AREA))
    import lup_colors as LC
    i = 4
    for v, inf in colors["info"].items():
        hx = inf["hex"]
        e = inf.get("lib_entry")
        vals = [v, "", "#" + hx, LC.hex_to_rgb_text(hx), LC.source_text(inf),
                f"{e['kode_sz'] or e['kode_z']} {e['subzona'] or e['zona']} ({e['hex']})" if e else "-",
                cnt[v], area[v]]
        for j, val in enumerate(vals, start=1):
            c = lg.cell(i, j, val)
            c.border = BOX
        colored_label(lg.cell(i, 1), hx)
        lg.cell(i, 2).fill = fill(hx)
        lg.cell(i, 8).number_format = FMT_SQM
        i += 1
    if colors["field"] != F_SUBZ and colors["groups"]:
        i += 1
        lg.cell(i, 1, f"{F_SUBZ} rows on pivots / slides (colour of their main {colors['field']})").font = Font(bold=True)
        i += 1
        header(lg, i, 1, [F_SUBZ, "Colour", "HEX"])
        for k, hx in colors["groups"].items():
            i += 1
            colored_label(lg.cell(i, 1, k), hx)
            lg.cell(i, 2).fill = fill(hx)
            lg.cell(i, 3, "#" + hx)
    for c, w in zip("ABCDEFGH", [30, 9, 10, 13, 26, 48, 10, 14]):
        lg.column_dimensions[c].width = w

    # checks
    ck = wb.create_sheet("Checks")
    header(ck, 1, 1, ["Check", "Result"])
    issues = data_checks(rows, S)
    for i, (k, v) in enumerate(issues, start=2):
        ck.cell(i, 1, k)
        ck.cell(i, 2, v)
    r0 = len(issues) + 4
    ck.cell(r0 - 1, 1, "Area check per polygon (attribute vs polygon geometry)").font = Font(bold=True)
    header(ck, r0, 1, ["Parcel / Layer", "Area_sqm (attribute)", "Polygon area (sqm)", "Difference"])
    for k, r in enumerate(rows, start=r0 + 1):
        ck.cell(k, 1, clean(r.get(F_PARCEL)) or clean(r.get("Layer")))
        ck.cell(k, 2, num(r.get(F_AREA))).number_format = "#,##0.00"
        ck.cell(k, 3, round(r["_geom_area"], 2)).number_format = "#,##0.00"
        ck.cell(k, 4, f"=C{k}-B{k}").number_format = "#,##0.00"
    for c, w in zip("ABCD", [45, 60, 20, 14]):
        ck.column_dimensions[c].width = w
    return wb, issues, real


def write_excel(path, fields, rows, S, site_total, colors, pivots=True, log=print):
    """Source sheet = raw attributes; summary sheets = Excel pivot tables on the Source
    table (built by Excel itself). Without Excel: SUMIFS formula sheets instead."""
    if pivots:
        wb, issues, real = _base_workbook(fields, rows, S, colors)
        for k, name in enumerate(PIVOT_SHEETS):
            wb.create_sheet(name, k)
        wb.active = 0
        wb.save(path)
        try:
            build_pivots(path, real, S, site_total, colors, log)
            return issues
        except Exception as e:  # no Excel / COM problem -> formula version
            log(f"Pivot tables need Microsoft Excel ({type(e).__name__}: {e}). Writing formula sheets instead.")
    wb, issues, real = _base_workbook(fields, rows, S, colors)
    _formula_sheets(wb, real, rows, S, site_total)
    for k, name in enumerate(["Overall MP", "Phasing", "Intensitas"]):
        wb.move_sheet(name, k - wb.sheetnames.index(name))
    wb.active = 0
    wb.save(path)
    return issues


def build_pivots(path, fields, S, site_total, colors, log=print):
    """Open the workbook in Excel (hidden) and add pivot tables on the Source table."""
    import comtypes
    import comtypes.client
    try:
        comtypes.CoInitialize()
    except OSError:
        pass
    import gc
    xl = comtypes.client.CreateObject("Excel.Application", dynamic=True)
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.ScreenUpdating = False
        _pivots_in_excel(xl, path, fields, S, site_total, colors, log)
    finally:
        gc.collect()  # release every workbook/pivot proxy while Excel is still running
        xl.Quit()
        del xl
        gc.collect()


def _pivots_in_excel(xl, path, fields, S, site_total, colors, log):
    wb = None
    try:
        # pivot field number formats are read in the local notation
        dec, th = xl.International(3), xl.International(4)
        F_INT, F_2, F_PCT = f"#{th}##0", f"#{th}##0{dec}00", f"0{dec}0%"
        wb = xl.Workbooks.Open(os.path.abspath(path))
        pc = wb.PivotCaches().Create(1, SOURCE_TABLE, 5)  # xlDatabase, table -> grows with the data
        have = set(fields)
        SUM, MAX, PCT_TOTAL, PCT_COL = -4157, -4136, 8, 7

        def bgr(hx):
            return int(hx[0:2], 16) + int(hx[2:4], 16) * 256 + int(hx[4:6], 16) * 65536

        def colour_column(ws, pt, field, scheme):
            if field not in have or not scheme:
                return
            try:
                c = pt.PivotFields(field).DataRange.Column
            except Exception:
                return
            col = ws.Columns(c)
            for value, hx in scheme.items():
                fc = col.FormatConditions.Add(1, 3, '="' + str(value).replace('"', '""') + '"')
                fc.Interior.Color = bgr(hx)
                fc.Font.Color = bgr("FFFFFF" if is_dark(hx) else "000000")
                fc.StopIfTrue = False

        def order_items(pt, field, order):
            pos = 1
            for name in order:
                try:
                    pt.PivotFields(field).PivotItems(name).Position = pos
                    pos += 1
                except Exception:
                    pass

        def make(sheet, name, title, rows_, data, page=None, subtotal_fields=()):
            ws = wb.Worksheets(sheet)
            ws.Range("A1").Value2 = title
            ws.Range("A1").Font.Bold = True
            ws.Range("A1").Font.Size = 13
            pt = pc.CreatePivotTable(ws.Range("A4"), name)
            pt.RowAxisLayout(1)  # tabular
            pt.TableStyle2 = "PivotStyleLight16"
            for i, f in enumerate(rows_):
                pf = pt.PivotFields(f)
                pf.Orientation = 1
                pf.Position = i + 1
            if page and page[0] in have:
                pf = pt.PivotFields(page[0])
                pf.Orientation = 3
                try:
                    pf.CurrentPage = page[1]
                except Exception:
                    pass
            for field, caption, func, calc, fmt in data:
                if field not in have:
                    continue
                df = pt.AddDataField(pt.PivotFields(field), caption, func)
                if calc:
                    df.Calculation = calc
                df.NumberFormat = fmt
            for f in rows_:
                if f not in subtotal_fields:
                    try:
                        pt.PivotFields(f).Subtotals = (False,) * 12
                    except Exception:
                        pass
            for f in (F_TYPE, F_SUBZ, F_INDT):
                if f in rows_:
                    colour_column(ws, pt, f, {F_TYPE: GROUP_COLORS, F_SUBZ: colors["groups"],
                                              F_INDT: {k: color_of(k) for k in S["ind"]}}[f])
            if colors["field"] in rows_ and colors["field"] not in (F_TYPE, F_SUBZ, F_INDT):
                colour_column(ws, pt, colors["field"], colors["scheme"])
            if F_TYPE in rows_:
                order_items(pt, F_TYPE, [SALEABLE, NON_SALEABLE])
            if F_SUBZ in rows_:
                order_items(pt, F_SUBZ, SALEABLE_ORDER + NON_SALEABLE_ORDER)
            ws.Range("A3:N3000").Columns.AutoFit()  # not row 1: the title would widen column A
            return ws, pt

        area = [(F_AREA, "Area (sqm)", SUM, None, F_INT),
                ("Area_Ha", "Area (Ha)", SUM, None, F_2),
                (F_AREA, "% of total", SUM, PCT_TOTAL, F_PCT)]
        ws, pt = make("Overall MP", "ptOverall", "OVERALL MASTERPLAN - LAND USE (pivot of Source)",
                      [F_TYPE, F_SUBZ], area, subtotal_fields=(F_TYPE,))
        # site total block (Road plug used on the slides)
        c0 = pt.TableRange2.Column + pt.TableRange2.Columns.Count + 1
        cells = [("Site total (sqm)", site_total or "=" + ws.Cells(6, c0 + 1).Address()),
                 ("Total in data (sqm)", '=GETPIVOTDATA("Area (sqm)",$A$4)'),
                 ("Adjustment added to Road on slides", f"={ws.Cells(5, c0 + 1).Address()}-{ws.Cells(6, c0 + 1).Address()}")]
        for k, (lab, val) in enumerate(cells):
            ws.Cells(5 + k, c0).Value2 = lab
            if isinstance(val, str):
                ws.Cells(5 + k, c0 + 1).Formula = val
            else:
                ws.Cells(5 + k, c0 + 1).Value2 = val
            ws.Cells(5 + k, c0 + 1).NumberFormat = F_2 if k == 2 else F_INT
        ws.Cells(4, c0).Value2 = "Site total check"
        ws.Cells(4, c0).Font.Bold = True
        ws.Columns(c0).AutoFit()
        ws.Columns(c0 + 1).ColumnWidth = 14

        if F_INDT in have:
            make("Industrial", "ptIndustrial", "INDUSTRIAL BREAKDOWN (pivot of Source, filtered to Industrial)",
                 [F_INDT], [(F_AREA, "Area (sqm)", SUM, None, F_INT), ("Area_Ha", "Area (Ha)", SUM, None, F_2),
                            (F_AREA, "% of Industrial", SUM, PCT_COL, F_PCT)], page=(F_SUBZ, "Industrial"))
        if F_PHASE in have:
            make("Phasing", "ptPhasing", "PHASING (pivot of Source)", [F_PHASE, F_SUBZ], area,
                 subtotal_fields=(F_PHASE,))
        if F_PARCEL in have:
            make("Intensitas", "ptIntensitas", "INTENSITAS - parcel intensity (pivot of Source)",
                 [F_TYPE, F_SUBZ, F_PARCEL],
                 [("BCR (%)", "BCR (%) ", MAX, None, "0"), ("FAR", "FAR ", MAX, None, f"0{dec}0#"),
                  ("Max Height", "Max Height ", MAX, None, "0"), ("GOSR (%)", "GOSR (%) ", MAX, None, "0"),
                  (F_AREA, "Area (sqm)", SUM, None, F_INT), ("GFA_sqm", "GFA (sqm)", SUM, None, F_INT),
                  (F_POP, "Population", SUM, None, F_INT)],
                 subtotal_fields=(F_TYPE,))
        if F_POPTYPE in have and F_POP in have:
            make("Population", "ptPopulation", "POPULATION (pivot of Source)", [F_POPTYPE],
                 [(F_POP, "Population", SUM, None, F_INT)])
        for name in PIVOT_SHEETS:  # drop sheets whose pivot could not be built
            ws = wb.Worksheets(name)
            if ws.PivotTables().Count == 0:
                ws.Delete()
        wb.Worksheets(1).Activate()
        wb.Save()
        log("Pivot tables built in Excel (refresh with Data > Refresh All after editing Source).")
    finally:
        if wb is not None:
            wb.Close(False)


def _formula_sheets(wb, fields, rows, S, site_total):
    n = len(rows) + 1

    def rng(field):  # missing optional field -> an empty column (SUMIFS then finds nothing)
        j = fields.index(field) + 1 if field in fields else len(fields) + 2
        L = get_column_letter(j)
        return f"Source!${L}$2:${L}${n}"

    AREA, TYPE, SUBZ, INDT = rng(F_AREA), rng(F_TYPE), rng(F_SUBZ), rng(F_INDT)
    PHASE = rng(F_PHASE)

    # ---- Overall MP -----------------------------------------------------
    mp = wb.create_sheet("Overall MP")
    mp["A1"] = "OVERALL MASTERPLAN - LAND USE SUMMARY"
    mp["A1"].font = Font(bold=True, size=14)
    header(mp, 3, 1, ["Land Use", "Area (sqm)", "Area (Ha)", "%"])
    r = 4
    total_row_placeholder = []

    def write_line(row, label, area_formula, hex_=None, bold=False, label_fill=None):
        c = mp.cell(row, 1, label)
        if hex_:
            colored_label(c, hex_, bold)
        else:
            c.font = Font(bold=bold)
            c.border = BOX
        mp.cell(row, 2, area_formula).number_format = FMT_SQM
        mp.cell(row, 3, f"=B{row}/10000").number_format = FMT_HA
        mp.cell(row, 4, "=B{0}/$B$TOTAL".format(row)).number_format = FMT_PCT
        total_row_placeholder.append(mp.cell(row, 4))
        for col in (2, 3, 4):
            mp.cell(row, col).border = BOX
            mp.cell(row, col).font = Font(bold=bold, color="FFFFFF" if label_fill and is_dark(label_fill) else "000000")
            if label_fill:
                mp.cell(row, col).fill = fill(label_fill)

    sale_hdr = r
    sale_rows = list(range(r + 1, r + 1 + len(S["sale"])))
    write_line(r, "SALEABLE", f"=SUM(B{sale_rows[0]}:B{sale_rows[-1]})", SALEABLE_HDR, True, SALEABLE_HDR)
    for row, s in zip(sale_rows, S["sale"]):
        write_line(row, s, f'=SUMIFS({AREA},{TYPE},"{SALEABLE}",{SUBZ},"{s}")', color_of(s))
    r = sale_rows[-1] + 2

    nons_hdr = r
    nons_rows = list(range(r + 1, r + 1 + len(S["nons"])))
    write_line(r, "NON-SALEABLE", f"=SUM(B{nons_rows[0]}:B{nons_rows[-1]})", NON_SALEABLE_HDR, True, NON_SALEABLE_HDR)
    road_row = None
    for row, s in zip(nons_rows, S["nons"]):
        f = f'=SUMIFS({AREA},{TYPE},"{NON_SALEABLE}",{SUBZ},"{s}")'
        if s == ROAD:
            road_row = row
            f += "+$B$ADJ"
        write_line(row, s, f, color_of(s))
    r = nons_rows[-1] + 2

    total_row = r
    data_sum = f"SUMIFS({AREA},{TYPE},\"{SALEABLE}\")+SUMIFS({AREA},{TYPE},\"{NON_SALEABLE}\")"
    write_line(r, "GRAND TOTAL", site_total if site_total else f"={data_sum}", "000000", True, "000000")
    r += 2
    mp.cell(r, 1, "Site total (input)").font = Font(italic=True)
    mp.cell(r, 2, site_total if site_total else "(sum of data)")
    mp.cell(r + 1, 1, "Total in shapefile").font = Font(italic=True)
    mp.cell(r + 1, 2, f"={data_sum}").number_format = FMT_SQM
    mp.cell(r + 2, 1, "Adjustment added to Road").font = Font(italic=True)
    adj_row = r + 2
    mp.cell(adj_row, 2, f"=B{total_row}-B{r + 1}").number_format = "#,##0.00"
    mp.cell(r + 3, 1, "Records with other Type").font = Font(italic=True)
    mp.cell(r + 3, 2, len(S["other"]))
    for row in mp.iter_rows():
        for c in row:
            if isinstance(c.value, str):
                c.value = c.value.replace("$TOTAL", f"${total_row}").replace("$ADJ", f"${adj_row}")

    # Industrial breakdown
    header(mp, 3, 6, ["Industrial Type", "Area (sqm)", "Area (Ha)", "% of Industrial", "% of Overall"])
    ind_rows = list(range(4, 4 + len(S["ind"])))
    ind_total = ind_rows[-1] + 1 if ind_rows else 4
    for row, t in zip(ind_rows, S["ind"]):
        colored_label(mp.cell(row, 6, t), color_of(t))
        crit = f'"{t}"' if t != "(unspecified)" else '""'
        mp.cell(row, 7, f'=SUMIFS({AREA},{SUBZ},"Industrial",{INDT},{crit})').number_format = FMT_SQM
        mp.cell(row, 8, f"=G{row}/10000").number_format = FMT_HA
        mp.cell(row, 9, f"=G{row}/$G${ind_total}").number_format = FMT_PCT
        mp.cell(row, 10, f"=G{row}/$B${total_row}").number_format = FMT_PCT
    colored_label(mp.cell(ind_total, 6, "Industrial Total"), color_of("Industrial"), True)
    if ind_rows:
        mp.cell(ind_total, 7, f"=SUM(G{ind_rows[0]}:G{ind_rows[-1]})").number_format = FMT_SQM
    mp.cell(ind_total, 8, f"=G{ind_total}/10000").number_format = FMT_HA
    mp.cell(ind_total, 9, f"=G{ind_total}/$G${ind_total}").number_format = FMT_PCT
    mp.cell(ind_total, 10, f"=G{ind_total}/$B${total_row}").number_format = FMT_PCT
    for row in range(4, ind_total + 1):
        for col in range(7, 11):
            mp.cell(row, col).border = BOX
            mp.cell(row, col).font = Font(bold=row == ind_total)

    # Population
    p0 = ind_total + 3
    header(mp, p0, 6, ["Population Type", "Population"])
    POPT, POP = rng(F_POPTYPE), rng(F_POP)
    prow = p0
    for prow, t in enumerate(S["pop"], start=p0 + 1):
        mp.cell(prow, 6, t).border = BOX
        c = mp.cell(prow, 7, f'=SUMIFS({POP},{POPT},"{t}")')
        c.number_format, c.border = FMT_SQM, BOX
    mp.cell(prow + 1, 6, "Grand Total").font = Font(bold=True)
    c = mp.cell(prow + 1, 7, f"=SUM(G{p0 + 1}:G{prow})")
    c.number_format, c.font = FMT_SQM, Font(bold=True)

    for col, w in zip("ABCDEFGHIJ", [22, 14, 12, 9, 3, 30, 14, 12, 15, 13]):
        mp.column_dimensions[col].width = w

    # ---- Phasing --------------------------------------------------------
    ph = wb.create_sheet("Phasing")
    phases = sorted({clean(x.get(F_PHASE)) or "(No phase)" for x in rows})
    last_phase = [p for p in phases if p != "(No phase)"][-1] if phases else None
    header(ph, 1, 1, ["Phase", "SUBZONA", "Area (sqm)", "Area (Ha)", "% of Site"])
    header(ph, 1, 8, ["Phase Summary", "Area (sqm)", "Area (Ha)", "% of Site"])
    r, sr = 2, 2
    tot_ref = f"'Overall MP'!$B${total_row}"
    adj_ref = f"'Overall MP'!$B${adj_row}"
    for p in phases:
        subs = ordered({clean(x.get(F_SUBZ)) for x in rows
                        if (clean(x.get(F_PHASE)) or "(No phase)") == p}, CHART_ORDER)
        if p == last_phase and ROAD not in subs:
            subs.append(ROAD)
        first = r
        pcrit = f'"{p}"' if p != "(No phase)" else '""'
        rowmap = {}
        for s in subs:
            ph.cell(r, 1, p if r == first else None).font = Font(bold=True)
            colored_label(ph.cell(r, 2, s), color_of(s))
            f = f'=SUMIFS({AREA},{PHASE},{pcrit},{SUBZ},"{s}")'
            if s == ROAD and p == last_phase:
                f += f"+{adj_ref}"
            ph.cell(r, 3, f).number_format = FMT_SQM
            ph.cell(r, 4, f"=C{r}/10000").number_format = FMT_HA
            ph.cell(r, 5, f"=C{r}/{tot_ref}").number_format = FMT_PCT
            rowmap[s] = r
            r += 1
        ph.cell(r, 1, f"{p} Total").font = Font(bold=True)
        for col, f in ((3, f"=SUM(C{first}:C{r - 1})"), (4, f"=C{r}/10000"), (5, f"=C{r}/{tot_ref}")):
            c = ph.cell(r, col, f)
            c.font, c.fill = Font(bold=True), fill(SUBTOTAL_GREY)
            c.number_format = [FMT_SQM, FMT_HA, FMT_PCT][col - 3]
        ph.cell(r, 1).fill = ph.cell(r, 2).fill = fill(SUBTOTAL_GREY)
        ptot = r
        r += 2
        # summary block
        c = ph.cell(sr, 8, p.upper())
        c.font, c.fill = Font(bold=True), fill(SUBTOTAL_GREY)
        for col, f in ((9, f"=C{ptot}"), (10, f"=D{ptot}"), (11, f"=C{ptot}/{tot_ref}")):
            c = ph.cell(sr, col, f)
            c.font, c.fill = Font(bold=True), fill(SUBTOTAL_GREY)
            c.number_format = [FMT_SQM, FMT_HA, FMT_PCT][col - 9]
        g = f"C{rowmap['Green']}" if "Green" in rowmap else "0"
        rd = f"C{rowmap[ROAD]}" if ROAD in rowmap else "0"
        for k, (lab, f) in enumerate((("Green", f"={g}"), ("Parcel", f"=I{sr}-(I{sr + 1}+I{sr + 3})"), ("Road", f"={rd}")), 1):
            ph.cell(sr + k, 8, lab)
            ph.cell(sr + k, 9, f).number_format = FMT_SQM
            ph.cell(sr + k, 10, f"=I{sr + k}/10000").number_format = FMT_HA
            ph.cell(sr + k, 11, f"=I{sr + k}/{tot_ref}").number_format = FMT_PCT
        sr += 5
    ph.cell(r, 1, "Grand Total").font = Font(bold=True)
    tot_cells = ",".join(f"C{x.row}" for x in ph["A"] if isinstance(x.value, str) and x.value.endswith(" Total"))
    c = ph.cell(r, 3, f"=SUM({tot_cells})")
    c.number_format, c.font = FMT_SQM, Font(bold=True)
    c = ph.cell(r, 4, f"=C{r}/10000")
    c.number_format, c.font = FMT_HA, Font(bold=True)
    for col, w in zip("ABCDEFGHIJK", [14, 18, 14, 12, 10, 3, 3, 16, 14, 12, 10]):
        ph.column_dimensions[col].width = w
    ph.freeze_panes = "A2"

    # ---- Intensitas -----------------------------------------------------
    it = wb.create_sheet("Intensitas")
    cols = ["Type", "SUBZONA", "Parcel_ID", "Phase", "BCR (%)", "FAR", "Max Height",
            "GOSR (%)", "Area (sqm)", "Area (Ha)", "GFA (sqm)", "Max Footprint (sqm)", "Pop_Total"]
    header(it, 1, 1, cols)
    parcels = defaultdict(list)
    for x in rows:
        if clean(x.get(F_PARCEL)):
            parcels[(clean(x.get(F_TYPE)), clean(x.get(F_SUBZ)))].append(x)
    r = 2
    for t in (SALEABLE, NON_SALEABLE):
        subs = ordered({s for (tt, s) in parcels if tt == t},
                       SALEABLE_ORDER if t == SALEABLE else NON_SALEABLE_ORDER)
        t_first = r
        for s in subs:
            s_first = r
            for x in sorted(parcels[(t, s)], key=lambda q: str(q.get(F_PARCEL))):
                it.cell(r, 1, t if r == t_first else None).font = Font(bold=True)
                if r == s_first:
                    colored_label(it.cell(r, 2, s), color_of(s))
                it.cell(r, 3, clean(x.get(F_PARCEL)))
                it.cell(r, 4, clean(x.get(F_PHASE)))
                for col, f in ((5, "BCR (%)"), (6, "FAR"), (7, "Max Height"), (8, "GOSR (%)")):
                    v = x.get(f)
                    it.cell(r, col, v if v not in (None, 0, "") else None)
                it.cell(r, 9, num(x.get(F_AREA))).number_format = FMT_SQM
                it.cell(r, 10, f"=I{r}/10000").number_format = "0.0000"
                it.cell(r, 11, f'=IF(F{r}="","",F{r}*I{r})').number_format = FMT_SQM
                it.cell(r, 12, f'=IF(E{r}="","",E{r}/100*I{r})').number_format = FMT_SQM
                it.cell(r, 13, num(x.get(F_POP)) or None).number_format = FMT_SQM
                r += 1
        it.cell(r, 1, f"{t} Total")
        for col, f in ((9, f"=SUM(I{t_first}:I{r - 1})"), (10, f"=I{r}/10000"),
                       (11, f"=SUM(K{t_first}:K{r - 1})"), (12, f"=SUM(L{t_first}:L{r - 1})"),
                       (13, f"=SUM(M{t_first}:M{r - 1})")):
            it.cell(r, col, f).number_format = FMT_SQM if col != 10 else FMT_HA
        for col in range(1, 14):
            it.cell(r, col).font = Font(bold=True)
            it.cell(r, col).fill = fill(SUBTOTAL_GREY)
        r += 2
    for col, w in zip("ABCDEFGHIJKLM", [14, 16, 11, 10, 9, 7, 11, 10, 13, 10, 13, 18, 11]):
        it.column_dimensions[col].width = w
    it.freeze_panes = "A2"


def data_checks(rows, S):
    out = [("Polygons read", len(rows)),
           ("Sum of Area_sqm in shapefile", round(S["shp_total"], 2)),
           ("Site total used for GRAND TOTAL", round(S["total"], 2)),
           ("Difference added to Road", round(S["road_adj"], 2))]
    diffs = [(clean(r.get(F_PARCEL)) or clean(r.get("Layer")), num(r.get(F_AREA)), r["_geom_area"])
             for r in rows]
    bad = [f"{p}: attr {a:,.0f} vs geom {g:,.0f}" for p, a, g in diffs if abs(a - g) > max(1.0, 0.001 * g)]
    out.append(("Area_sqm differs from polygon area (>0.1%)", "; ".join(bad) or "OK"))
    no_type = [clean(r.get("Layer")) for r in rows if clean(r.get(F_TYPE)) not in (SALEABLE, NON_SALEABLE)]
    out.append(("Records without Saleable/Non-Saleable Type", ", ".join(map(str, no_type)) or "OK"))
    no_sub = [clean(r.get("Layer")) for r in rows if not clean(r.get(F_SUBZ))]
    out.append(("Records without SUBZONA", ", ".join(map(str, no_sub)) or "OK"))
    no_ph = [clean(r.get(F_PARCEL)) or clean(r.get("Layer")) for r in rows if not clean(r.get(F_PHASE))]
    out.append(("Records without Phase", ", ".join(map(str, no_ph)) or "OK"))
    seen, dup = set(), set()
    for r in rows:
        p = clean(r.get(F_PARCEL))
        if p:
            (dup if p in seen else seen).add(p)
    out.append(("Duplicate Parcel_ID (areas are summed)", ", ".join(sorted(dup)) or "OK"))
    ind_no = [clean(r.get(F_PARCEL)) for r in rows
              if clean(r.get(F_SUBZ)) == "Industrial" and not clean(r.get(F_INDT))]
    out.append(("Industrial parcels without IND_TYPE", ", ".join(map(str, ind_no)) or "OK"))
    gfa_bad = []
    for r in rows:
        far, gfa, a = r.get("FAR"), r.get("GFA_sqm"), num(r.get(F_AREA))
        if far and gfa and a and abs(num(gfa) - num(far) * a) > max(10.0, 0.01 * num(far) * a):
            gfa_bad.append(f"{clean(r.get(F_PARCEL)) or clean(r.get('Layer'))}: {num(gfa):,.0f} vs {num(far) * a:,.0f}")
    out.append(("GFA_sqm differs from FAR x Area_sqm (>1%)", "; ".join(gfa_bad) or "OK"))
    return out


# --------------------------------------------------------------------------
# 4. POWERPOINT
# --------------------------------------------------------------------------
def fmt_int(v):
    s = f"{v:,.0f}"
    return s.replace(",", ".") if NUMBER_STYLE == "id" else s


def fmt_dec(v, d=2):
    s = f"{v:,.{d}f}"
    return s.translate(str.maketrans(",.", ".,")) if NUMBER_STYLE == "id" else s


def fmt_pct(v, d=1):
    return fmt_dec(v * 100, d) + "%"


def rgb(hex_):
    return RGBColor.from_string(hex_)


def set_cell_border(cell, sides, color="D9D9D9", width=Pt(0.75)):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for e in tcPr.findall(qn(tag)):
            tcPr.remove(e)
    for i, (tag, key) in enumerate((("a:lnL", "L"), ("a:lnR", "R"), ("a:lnT", "T"), ("a:lnB", "B"))):
        ln = etree.SubElement(tcPr, qn(tag), w=str(int(width)))
        if key in sides:
            sf = etree.SubElement(ln, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr"), val=sides[key] if isinstance(sides, dict) else color)
        else:
            etree.SubElement(ln, qn("a:noFill"))
        tcPr.remove(ln)
        tcPr.insert(i, ln)


def style_cell(cell, text, size=6, bold=False, fill_hex=None, color=None,
               align=PP_ALIGN.LEFT, indent=0, bottom_line=None):
    cell.text = ""
    tf = cell.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size, run.font.bold = Pt(size), bold
    if color is None:
        color = "FFFFFF" if fill_hex and is_dark(fill_hex) else "000000"
    run.font.color.rgb = rgb(color)
    # empty cells otherwise inherit an 18pt line height and inflate the row
    p._p.get_or_add_endParaRPr().set("sz", str(int(size * 100)))
    cell.margin_left = Emu(36000 + indent)
    cell.margin_right, cell.margin_top, cell.margin_bottom = Emu(36000), Emu(0), Emu(0)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if fill_hex:
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(fill_hex)
    else:
        cell.fill.background()
    set_cell_border(cell, {"B": bottom_line} if bottom_line else {})


def new_table(slide, nrows, col_widths, left, top, row_h):
    gf = slide.shapes.add_table(nrows, len(col_widths), Emu(left), Emu(top),
                                Emu(sum(col_widths)), Emu(row_h * nrows))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
    sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # No Style, No Grid
    for i, w in enumerate(col_widths):
        tbl.columns[i].width = Emu(w)
    for r in tbl.rows:
        r.height = Emu(row_h)
    return gf, tbl


def full_plot_area(chart):
    """Make the plot fill the whole chart frame so stacked charts line up exactly."""
    pa = chart._chartSpace.find(".//" + qn("c:plotArea"))
    lay = pa.find(qn("c:layout"))
    if lay is None:
        lay = etree.Element(qn("c:layout"))
        pa.insert(0, lay)
    for ch in list(lay):
        lay.remove(ch)
    ml = etree.SubElement(lay, qn("c:manualLayout"))
    for tag, val in (("layoutTarget", "inner"), ("xMode", "edge"), ("yMode", "edge"),
                     ("x", "0"), ("y", "0"), ("w", "1"), ("h", "1")):
        etree.SubElement(ml, qn("c:" + tag), val=val)


def add_ring_chart(slide, kind, cats, vals, colors, cx, cy, diameter, hole=None, line="FFFFFF"):
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("Area", vals)
    gf = slide.shapes.add_chart(kind, Emu(cx - diameter // 2), Emu(cy - diameter // 2),
                                Emu(diameter), Emu(diameter), cd)
    ch = gf.chart
    ch.has_title = False
    ch.has_legend = False
    full_plot_area(ch)
    plot = ch.plots[0]
    plot.vary_by_categories = True
    if hole is not None:
        hs = ch._chartSpace.find(".//" + qn("c:holeSize"))
        if hs is None:
            dn = ch._chartSpace.find(".//" + qn("c:doughnutChart"))
            hs = etree.SubElement(dn, qn("c:holeSize"))
        hs.set("val", str(max(10, min(90, int(round(hole))))))
    for i, pt in enumerate(plot.series[0].points):
        if colors[i] is None:
            pt.format.fill.background()
            pt.format.line.fill.background()
            continue
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = rgb(colors[i])
        # white separators, but outline white slices (Road) so they stay visible
        pt.format.line.color.rgb = rgb("BFBFBF" if colors[i].upper() == "FFFFFF" else line)
        pt.format.line.width = Pt(0.75)
    return gf


def polar(cx, cy, r, frac):
    a = frac * 2 * math.pi
    return cx + r * math.sin(a), cy - r * math.cos(a)


_ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}


def add_text(slide, x, y, w, h, text, size=7, bold=False, align="l", color="000000", anchor="m"):
    tb = slide.shapes.add_textbox(Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.TOP if anchor == "t" else MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
    for k, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.alignment = _ALIGN[align]
        r = p.add_run()
        r.text = line
        r.font.size, r.font.bold = Pt(size), bold if k == 0 else False
        r.font.color.rgb = rgb(color)
    return tb


# ---- slide scene --------------------------------------------------------
# Each slide is described as a list of simple items in EMU (914400 per inch).
# write_pptx() turns them into PowerPoint shapes; lup_widget.py paints the same
# items on screen, so the preview always matches the exported file.
#   text  : x y w h text size bold color align(l/c/r) anchor(t/m)   first line bold
#   table : left top widths rows=[{h, cells=[cell|None]}]   None = covered by span
#           cell = {text size bold fill color align indent bline span}
#   chart : kind(pie/doughnut) cats vals colors(None = hidden) cx cy d hole(%)
#   line  : x0 y0 x1 y1 color w(pt)
#   rect  : x y w h color w_pt dash
SLIDE_W, SLIDE_H = 9144000, 6858000


def T_(x, y, w, h, text, size=7, bold=False, color="000000", align="l", anchor="m"):
    return dict(t="text", x=x, y=y, w=w, h=h, text=text, size=size, bold=bold,
                color=color, align=align, anchor=anchor)


def C_(text="", size=6, bold=False, fill=None, color=None, align="l", indent=0, bline=None, span=1):
    if color is None:
        color = "FFFFFF" if fill and is_dark(fill) else "000000"
    return dict(text=text, size=size, bold=bold, fill=fill, color=color, align=align,
                indent=indent, bline=bline, span=span)


def slice_labels(vals, labels, colors, cx, cy, diameter, kind, hole=None, size=7):
    """Text items centred in each labelled slice (pie: 60% radius, ring: mid-ring)."""
    out, tot, cum = [], float(sum(vals)), 0.0
    r_lab = diameter / 2 * (0.6 if kind == "pie" else (1 + (hole or 50) / 100) / 2)
    for i, v in enumerate(vals):
        if labels[i] and colors[i]:
            x, y = polar(cx, cy, r_lab, (cum + v / 2) / tot)
            w, h = Pt(size * 6), Pt(size * 2.6)
            out.append(T_(x - w / 2, y - h / 2, w, h, labels[i], size, True,
                          "FFFFFF" if is_dark(colors[i]) else "000000", "c"))
        cum += v
    return out


def callouts(items, cx, cy, r_out, size=8):
    """items: (label, mid_frac) -> leader lines + labels outside the ring."""
    lab_h = Pt(size * 1.6)
    placed, out = [], []
    for label, frac in sorted(items, key=lambda t: t[1]):
        x0, y0 = polar(cx, cy, r_out, frac)
        x1, y1 = polar(cx, cy, r_out * 1.18, frac)
        right = x1 >= cx
        for py in placed:  # avoid overlaps
            if abs(py[0] - y1) < lab_h and py[1] == right:
                y1 = py[0] + lab_h * (1 if (frac % 1) < 0.5 else -1)
        placed.append((y1, right))
        out.append(dict(t="line", x0=x0, y0=y0, x1=x1, y1=y1, color="000000", w=0.75))
        w = Pt(size * 0.62 * max(len(s) for s in label.split("\n")) + 6)
        tx = x1 + Pt(2) if right else x1 - w - Pt(2)
        out.append(T_(tx, y1 - lab_h / 2, w, lab_h, label, size, True, "000000",
                      "l" if right else "r"))
    return out


def chart_series(S):
    allv = OrderedDict(list(S["nons"].items()) + list(S["sale"].items()))
    cats = [c for c in CHART_ORDER if c in allv] + [c for c in allv if c not in CHART_ORDER]
    cats = [c for c in cats if allv[c] > 0]
    return cats, [allv[c] for c in cats]


def pie_labels(cats, vals, total, min_share=0.035):
    return [f"{SHORT.get(c, c)}\n{fmt_pct(v / total)}" if v / total >= min_share else None
            for c, v in zip(cats, vals)]


def title_item(title):
    return T_(345591, 477838, 6000000, 900000, title, 40, False, "000000", "l", "t")


def scene_summary(S, title="Land Use Summary"):
    T = S["total"]
    st, nt = sum(S["sale"].values()), sum(S["nons"].values())
    big, small, gap = 173185, 124635, 135429
    widths = [640000, 450000, 410000, 330000]
    left, top = 6221586, 1766050

    def nums(v, **kw):
        return [C_(t, align="r", **kw) for t in (fmt_int(v), fmt_dec(v / 1e4), fmt_pct(v / T))]

    rows = [dict(h=big, cells=[C_(fill="000000")] +
                 [C_(t, 6, True, "000000", align="c") for t in ("Area (sqm)", "Area (Ha)", "%")])]
    for lab, v, hx, items in (("SALEABLE", st, SALEABLE_HDR, S["sale"]),
                              ("NON-SALEABLE", nt, NON_SALEABLE_HDR, S["nons"])):
        rows.append(dict(h=big, cells=[C_(lab, 6, True, hx, "FFFFFF")] +
                         nums(v, size=6, bold=True, fill=hx, color="FFFFFF")))
        for k, val in items.items():
            rows.append(dict(h=small, cells=[C_(k, 5, False, color_of(k), bline="FFFFFF")] +
                             nums(val, size=5, bline="BFBFBF")))
        rows.append(dict(h=gap, cells=[C_(size=5) for _ in widths]))
    rows.append(dict(h=big, cells=[C_("GRAND TOTAL", 6, True, "000000")] +
                     nums(T, size=6, bold=True, fill="000000")))
    items = [title_item(title), dict(t="table", left=left, top=top, widths=widths, rows=rows)]

    # pie inside a Saleable / Non-Saleable ring
    cats, vals = chart_series(S)
    d_out = 1_500_000
    cx = left + sum(widths) // 2
    cy = top + sum(r["h"] for r in rows) + 120_000 + d_out // 2
    d_in = int(d_out * 0.74)
    hole = d_in / d_out * 100
    cols = [color_of(c) for c in cats]
    items.append(dict(t="chart", kind="doughnut", cats=[NON_SALEABLE, SALEABLE], vals=[nt, st],
                      colors=[GROUP_COLORS[NON_SALEABLE], GROUP_COLORS[SALEABLE]],
                      cx=cx, cy=cy, d=d_out, hole=hole))
    items.append(dict(t="chart", kind="pie", cats=cats, vals=vals, colors=cols,
                      cx=cx, cy=cy, d=d_in, hole=None))
    items += slice_labels(vals, pie_labels(cats, vals, T, 0.10), cols, cx, cy, d_in, "pie", size=5)
    if T:
        items += callouts([(f"NON-SALEABLE {fmt_pct(nt / T, 0)}", nt / T / 2),
                           (f"SALEABLE {fmt_pct(st / T, 0)}", nt / T + st / T / 2)],
                          cx, cy, d_out / 2, size=6)
    return dict(title=title, items=items)


def scene_industrial(S, title="Land Use Program &\nIndustrial Breakdown"):
    T = S["total"]
    st, nt = sum(S["sale"].values()), sum(S["nons"].values())
    widths = [168198, 1685759, 633590, 552620, 413959]
    left, top, rh, gap_h = 1117874, 1766050, 170336, 102641

    def nums(v, **kw):
        return [C_(t, 6, align="r", **kw) for t in (fmt_int(v), fmt_dec(v / 1e4), fmt_pct(v / T))]

    def wide(text, **kw):  # label spanning the first two columns
        return [C_(text, span=2, **kw), None]

    rows = [dict(h=rh, cells=wide("PROGRAM", size=6, bold=True, fill="000000", align="c") +
                 [C_(t, 6, True, "000000", align="c") for t in ("Area (sqm)", "Area (Ha)", "%")])]
    sub_rows = []
    for k, v in S["sale"].items():
        rows.append(dict(h=rh, cells=wide(k, size=6, bold=k == "Industrial", fill=color_of(k),
                                          bline="FFFFFF") + nums(v, bline="BFBFBF")))
        if k == "Industrial":
            for t, a in S["ind"].items():
                hx = color_of(t)
                sub_rows.append(len(rows))
                rows.append(dict(h=rh, cells=[C_(fill=hx, bline="FFFFFF"),
                                              C_(t, 6, False, hx, indent=72000, bline="FFFFFF")] +
                                 nums(a, color="808080", bline="BFBFBF")))
    rows.append(dict(h=rh, cells=wide("SALEABLE", size=6, bold=True, fill=SUBTOTAL_GREY, align="r") +
                     nums(st, bold=True, fill=SUBTOTAL_GREY)))
    rows.append(dict(h=gap_h, cells=[C_(size=5) for _ in widths]))
    for k, v in S["nons"].items():
        rows.append(dict(h=rh, cells=wide(k, size=6, fill=color_of(k), bline="FFFFFF") +
                         nums(v, bline="BFBFBF")))
    rows.append(dict(h=rh, cells=wide("NON-SALEABLE", size=6, bold=True, fill=SUBTOTAL_GREY, align="r") +
                     nums(nt, bold=True, fill=SUBTOTAL_GREY)))
    rows.append(dict(h=gap_h, cells=[C_(size=5) for _ in widths]))
    rows.append(dict(h=rh, cells=wide("GRAND TOTAL", size=6, bold=True, fill="000000", align="r") +
                     nums(T, bold=True, fill="000000")))
    items = [title_item(title), dict(t="table", left=left, top=top, widths=widths, rows=rows)]

    if sub_rows:  # dashed box around the industrial sub-types
        y0 = top + sum(r["h"] for r in rows[:sub_rows[0]])
        h = sum(rows[k]["h"] for k in sub_rows)
        items.append(dict(t="rect", x=left - 20000, y=y0 - 15000, w=sum(widths) + 40000,
                          h=h + 30000, color="808080", w_pt=2, dash=True))

    # pie of all land uses + outer arc that explodes the Industrial slice
    cats, vals = chart_series(S)
    d_in, d_out = 2_000_000, 3_000_000
    cx, cy = 6_250_000, 3_750_000
    hole = d_in / d_out * 100
    cols = [color_of(c) for c in cats]
    if S["ind"] and "Industrial" in cats:
        before = sum(vals[:cats.index("Industrial")])
        ind_vals = list(S["ind"].values())
        rv = [before] + ind_vals + [max(0.0, T - before - sum(ind_vals))]
        rcol = [None] + [color_of(t) for t in S["ind"]] + [None]
        labels, outside, cum = [None], [], before
        for t, v in S["ind"].items():
            share = v / T
            if share >= 0.04:
                labels.append(f"{SHORT.get(t, t)}\n{fmt_pct(share)}")
            else:
                labels.append(None)
                outside.append((f"{SHORT.get(t, t)} {fmt_pct(share)}", (cum + v / 2) / T))
            cum += v
        labels.append(None)
        items.append(dict(t="chart", kind="doughnut", cats=["_before"] + list(S["ind"]) + ["_after"],
                          vals=rv, colors=rcol, cx=cx, cy=cy, d=d_out, hole=hole))
    items.append(dict(t="chart", kind="pie", cats=cats, vals=vals, colors=cols,
                      cx=cx, cy=cy, d=d_in, hole=None))
    items += slice_labels(vals, pie_labels(cats, vals, T, 0.05), cols, cx, cy, d_in, "pie", size=7)
    if S["ind"] and "Industrial" in cats:
        items += slice_labels(rv, labels, rcol, cx, cy, d_out, "doughnut", hole, size=8)
        items += callouts(outside, cx, cy, d_out / 2, size=8)
    return dict(title=title, items=items)


def build_slides(S):
    return [scene_summary(S), scene_industrial(S)]


def render_slide_pptx(prs, layout, scene):
    s = prs.slides.add_slide(layout)
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    for it in scene["items"]:
        k = it["t"]
        if k == "text":
            add_text(s, it["x"], it["y"], it["w"], it["h"], it["text"], it["size"], it["bold"],
                     it["align"], it["color"], it["anchor"])
        elif k == "table":
            rows = it["rows"]
            gf, tbl = new_table(s, len(rows), it["widths"], it["left"], it["top"], rows[0]["h"])
            for i, row in enumerate(rows):
                tbl.rows[i].height = Emu(row["h"])
                cells = list(tbl.rows[i].cells)
                for j, c in enumerate(row["cells"]):
                    if c is None:
                        continue
                    if c["span"] > 1:
                        cells[j].merge(cells[j + c["span"] - 1])
                    style_cell(cells[j], c["text"], c["size"], c["bold"], c["fill"], c["color"],
                               _ALIGN[c["align"]], c["indent"], c["bline"])
        elif k == "chart":
            kind = XL_CHART_TYPE.PIE if it["kind"] == "pie" else XL_CHART_TYPE.DOUGHNUT
            add_ring_chart(s, kind, it["cats"], it["vals"], it["colors"], it["cx"], it["cy"],
                           it["d"], it["hole"])
        elif k == "line":
            ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(int(it["x0"])), Emu(int(it["y0"])),
                                        Emu(int(it["x1"])), Emu(int(it["y1"])))
            ln.line.color.rgb = rgb(it["color"])
            ln.line.width = Pt(it["w"])
        elif k == "rect":
            box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(int(it["x"])), Emu(int(it["y"])),
                                     Emu(int(it["w"])), Emu(int(it["h"])))
            box.fill.background()
            box.line.color.rgb = rgb(it["color"])
            box.line.width = Pt(it["w_pt"])
            if it["dash"]:
                box.line.dash_style = MSO_LINE_DASH_STYLE.DASH
            box.shadow.inherit = False
    return s


def blank_deck(template):
    if template and os.path.exists(template):
        prs = Presentation(template)
        layout = (prs.slides[0].slide_layout if len(prs.slides)
                  else prs.slide_layouts[min(6, len(prs.slide_layouts) - 1)])
        sld_ids = prs.slides._sldIdLst
        for sld in list(sld_ids):
            prs.part.drop_rel(sld.rId)
            sld_ids.remove(sld)
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Emu(9144000), Emu(6858000)
        layout = prs.slide_layouts[6]
    return prs, layout


def write_pptx(path, S, template, slides=None):
    prs, layout = blank_deck(template)
    for scene in slides or build_slides(S):
        render_slide_pptx(prs, layout, scene)
    prs.save(path)


# --------------------------------------------------------------------------
def summary_lines(S):
    T = S["total"]
    out = ["Land use                Area (sqm)   Area (Ha)      %"]
    for grp, d in (("SALEABLE", S["sale"]), ("NON-SALEABLE", S["nons"])):
        g = sum(d.values())
        out.append(f"{grp:<22}{g:>12,.0f}{g / 1e4:>12,.2f}{g / T:>8.1%}")
        out += [f"  {k:<20}{v:>12,.0f}{v / 1e4:>12,.2f}{v / T:>8.1%}" for k, v in d.items()]
    out.append(f"{'GRAND TOTAL':<22}{T:>12,.0f}{T / 1e4:>12,.2f}{1:>8.1%}")
    return out


def make_colors(fields, rows, field=None, priority=None, library=None, layer_scheme=None, lib_dir=None):
    """Colour scheme for the outputs (see lup_colors). Arguments override lup_colours.json."""
    import lup_colors as LC
    cfg = LC.load_config(lib_dir or HERE)
    if priority:
        cfg["priority"] = priority
    if library:
        cfg["library"] = library
    lib = LC.load_library(LC.library_path(cfg, lib_dir or HERE))
    return LC.build(rows, fields, field or cfg.get("field") or F_SUBZ, cfg, lib, layer_scheme,
                    BASE_COLORS, F_SUBZ, F_AREA)


def run(fields, rows, xlsx, pptx, site_total=SITE_TOTAL_SQM, template=DEFAULT_TEMPLATE,
        number_style=None, log=print, colors=None, pivots=True):
    """Core pipeline, shared by the command line, the QGIS Processing tool and the window.

    fields: attribute names in source order.
    rows:   one dict per polygon (field -> value) plus "_geom_area" (sqm).
    colors: result of make_colors() / lup_colors.build(); None = from lup_colours.json.
    Either output path may be None to skip it. Returns (summary, checks).
    """
    global NUMBER_STYLE
    if number_style:
        NUMBER_STYLE = number_style
    site_total = site_total or None
    missing = [f for f in (F_TYPE, F_SUBZ, F_AREA) if f not in fields]
    if missing:
        raise ValueError("Layer is missing required field(s): " + ", ".join(missing))
    if colors is None:
        colors = make_colors(fields, rows)
    set_palette(colors["groups"])
    log(f"Colours: field '{colors['field']}', {len(colors['scheme'])} categories")
    for f in (F_INDT, F_PHASE, F_PARCEL, F_POP, F_POPTYPE):
        if f not in fields:  # optional fields: add empty so the sheets still build
            fields = fields + [f]
    S = summarise(rows, site_total)
    log(f"{len(rows)} polygons, {S['shp_total']:,.2f} sqm in data, grand total "
        f"{S['total']:,.0f} sqm (Road adjusted by {S['road_adj']:+,.2f})")
    issues = []
    for path, label, app, write in (
            (xlsx, "Excel", "Excel", lambda p: write_excel(p, fields, rows, S, site_total, colors, pivots, log)),
            (pptx, "PowerPoint", "PowerPoint", lambda p: write_pptx(p, S, template))):
        if not path:
            continue
        log(f"Writing {label}: {path}")
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        try:
            res = write(path)
        except PermissionError:
            raise RuntimeError(f"Cannot write {path} - close it in {app} and run again.")
        if label == "Excel":
            issues = res
    for k, v in issues[4:]:
        if v != "OK":
            log(f"CHECK  {k}: {v}")
    return S, issues


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--shp", default=DEFAULT_SHP, help="land use shapefile (.shp)")
    ap.add_argument("--out", default=DEFAULT_OUT, help="output folder")
    ap.add_argument("--template", default=DEFAULT_TEMPLATE, help="PowerPoint template (optional)")
    ap.add_argument("--site-total", type=float, default=SITE_TOTAL_SQM,
                    help="site area in sqm for GRAND TOTAL (0 = sum of shapefile)")
    ap.add_argument("--name", default=None, help="base name for outputs (default: shapefile name)")
    ap.add_argument("--color-field", default=None, help="field the colours are keyed on (default from lup_colours.json)")
    ap.add_argument("--priority", choices=["layer", "library", "palette"], default=None,
                    help="colour source order (default from lup_colours.json)")
    ap.add_argument("--library", default=None, help="colour library workbook (RDTR colour code)")
    ap.add_argument("--no-pivots", action="store_true", help="formula sheets instead of Excel pivot tables")
    a = ap.parse_args()

    site_total = a.site_total or None
    os.makedirs(a.out, exist_ok=True)
    base = a.name or os.path.splitext(os.path.basename(a.shp))[0]
    xlsx = os.path.join(a.out, f"{base}_LUP_Database.xlsx")
    pptx = os.path.join(a.out, f"{base}_LUP_Tables.pptx")

    print(f"Reading {a.shp}")
    fields, rows = read_shapefile(a.shp)
    try:
        colors = make_colors(fields, rows, a.color_field, a.priority, a.library)
        S, _ = run(fields, rows, xlsx, pptx, site_total, a.template,
                   log=lambda m: print("      " + m), colors=colors, pivots=not a.no_pivots)
    except (RuntimeError, ValueError) as e:
        sys.exit("      " + str(e))
    print()
    for line in summary_lines(S):
        print("      " + line)
    print("\nDone.")


if __name__ == "__main__":
    main()
