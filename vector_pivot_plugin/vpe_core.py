"""Vector Pivot Export - Excel writer (no QGIS imports, testable on its own).

Input is plain Python data:
    fields  : list of column names for the Source sheet
    rows    : list of dicts {field: value}
    spec    : dict(row_fields=[...], col_field=None|str,
                   measures=[(field, agg)], percent=bool)
              agg in "sum", "mean", "min", "max", "count"
    colors  : {field: {value: "RRGGBB"}}   (from the layer symbology)
    orders  : {field: [value, ...]}        (legend order; others sorted)

Output workbook:
    Pivot   - native Excel pivot table on the Source table (needs Excel on Windows),
              or a static pivot with the same layout when Excel is not available
    Source  - the attribute table as an Excel table (colour coded)
    Colours - legend: value, swatch, HEX, feature count per coloured field
"""

import os
from collections import Counter, defaultdict

import xlsxwriter

SOURCE_TABLE = "tblSource"
BLANK = "(blank)"
AGG_LABEL = {"sum": "Sum", "mean": "Average", "min": "Min", "max": "Max", "count": "Count"}
ADDITIVE = ("sum", "count")
HDR = "404040"
SUB = "F2F2F2"


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def is_dark(hx):
    r, g, b = int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) < 140


def key(v):
    return BLANK if v is None or (isinstance(v, str) and not v.strip()) else v


def rank(v, order=None):
    """Sort key: legend order first, then numbers, then text, (blank) last."""
    if order and v in order:
        return (0, order.index(v), 0, "")
    if v == BLANK:
        return (3, 0, 0, "")
    if isinstance(v, (int, float)):
        return (1, 0, v, "")
    return (2, 0, 0, str(v).lower())


def sort_values(values, order=None):
    return sorted(dict.fromkeys(values), key=lambda v: rank(v, order))


def aggregate(vals, agg):
    if agg == "count":
        return sum(1 for v in vals if v is not None)
    vals = [v for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]
    if not vals:
        return None
    if agg == "sum":
        return sum(vals)
    if agg == "mean":
        return sum(vals) / len(vals)
    return min(vals) if agg == "min" else max(vals)


def caption(field, agg):
    return f"{AGG_LABEL[agg]} of {field}"


def number_format(field, agg):
    if agg == "count" or field in ("Features",):
        return "#,##0"
    return "#,##0.00"


# ---------------------------------------------------------------------------
# static pivot (pure Python)
# ---------------------------------------------------------------------------
def compute_pivot(rows, spec, orders=None):
    """-> dict(row_keys, subtotal_of, col_values, cells, totals) with aggregated values.

    cells[(rowkey, col, m)] where rowkey is a tuple over row_fields, col a column value
    (or None when there is no column field) and m the measure index. Subtotals use the
    1-element prefix key, the grand total the empty tuple; col "__total__" = across columns.
    """
    orders = orders or {}
    rf, cf, ms = spec["row_fields"], spec.get("col_field"), spec["measures"]
    groups = defaultdict(lambda: defaultdict(list))  # (prefix, col) -> m -> values
    for r in rows:
        rk = tuple(key(r.get(f)) for f in rf)
        ck = key(r.get(cf)) if cf else None
        prefixes = [rk, ()] + ([rk[:1]] if len(rf) > 1 else [])
        cols = [ck, "__total__"] if cf else [ck]
        for p in prefixes:
            for c in cols:
                for i, (f, _) in enumerate(ms):
                    groups[(p, c)][i].append(r.get(f))
    row_keys = sorted({p for p, _ in groups if len(p) == len(rf) and p},
                      key=lambda t: tuple(rank(v, orders.get(f)) for v, f in zip(t, rf)))
    col_values = sort_values([c for _, c in groups if c not in (None, "__total__")], orders.get(cf)) if cf else [None]
    cells = {}
    for (p, c), mv in groups.items():
        for i, (_, agg) in enumerate(ms):
            cells[(p, c, i)] = aggregate(mv[i], agg)
    return dict(row_keys=row_keys, col_values=col_values, cells=cells)


class Formats:
    def __init__(self, wb):
        self.wb = wb
        self.cache = {}

    def get(self, **kw):
        k = tuple(sorted(kw.items()))
        if k not in self.cache:
            props = dict(border=1, border_color="BFBFBF", valign="vcenter")
            props.update(kw)
            self.cache[k] = self.wb.add_format(props)
        return self.cache[k]

    def label(self, hx=None, bold=False):
        if not hx:
            return self.get(bold=bold)
        return self.get(bold=bold, bg_color="#" + hx, font_color="#FFFFFF" if is_dark(hx) else "#000000")

    def header(self, hx=None):
        hx = hx or HDR
        return self.get(bold=True, text_wrap=True, align="center", bg_color="#" + hx,
                        font_color="#FFFFFF" if is_dark(hx) else "#000000")


def pivot_grid(rows, spec, colors, orders=None):
    """Layout of the pivot, shared by the static Excel sheet and the preview window.

    -> dict(row_fields, col_field, cols, blocks, lines, header_top, header_sub)
       blocks: [(caption, measure index, percent?, agg, field)]
       cols:   column values (+ "__total__" when there is a column field)
       header_top: [(caption, span)] per block; header_sub: [(text, hex)] per value column
       lines:  [(kind, labels, values)] kind "row" / "subtotal" / "total";
               labels [(text, hex)] per row field; values [(number|None, number format)]
    """
    rf, cf, ms = spec["row_fields"], spec.get("col_field"), spec["measures"]
    P = compute_pivot(rows, spec, orders)
    cols = P["col_values"] + (["__total__"] if cf else [])
    blocks = []
    for i, (f, agg) in enumerate(ms):
        blocks.append((caption(f, agg), i, False, agg, f))
        if spec.get("percent") and agg in ADDITIVE:
            blocks.append((f"% of total ({caption(f, agg)})", i, True, agg, f))
    header_top = [(b[0], len(cols)) for b in blocks]
    header_sub = []
    if cf:
        for _ in blocks:
            for cv in cols:
                header_sub.append(("Total", None) if cv == "__total__" else (cv, (colors.get(cf) or {}).get(cv)))

    def grand(i):
        return P["cells"].get(((), "__total__" if cf else None, i))

    def values(pkey):
        out = []
        for _, i, pct, agg, f in blocks:
            for cv in cols:
                v = P["cells"].get((pkey, cv, i))
                if pct and v is not None:
                    g = grand(i)
                    v = v / g if g else None
                out.append((v, "0.0%" if pct else number_format(f, agg)))
        return out

    def colour(f, v):
        return (colors.get(f) or {}).get(v)

    nl = len(rf)
    lines, prev = [], None

    def subtotal(v):
        lines.append(("subtotal", [(f"{v} Total", colour(rf[0], v))] + [("", None)] * (nl - 1), values((v,))))

    for rk in P["row_keys"]:
        if nl > 1 and prev is not None and rk[0] != prev:
            subtotal(prev)
        lines.append(("row", [(v, colour(f, v)) for f, v in zip(rf, rk)], values(rk)))
        prev = rk[0]
    if nl > 1 and prev is not None:
        subtotal(prev)
    lines.append(("total", [("Grand Total", None)] + [("", None)] * (nl - 1), values(())))
    return dict(row_fields=rf, col_field=cf, cols=cols, blocks=blocks, lines=lines,
                header_top=header_top, header_sub=header_sub)


def write_static_pivot(wb, ws, F, rows, spec, colors, orders, title):
    G = pivot_grid(rows, spec, colors, orders)
    rf, cf = G["row_fields"], G["col_field"]
    nl = len(rf)
    ws.write(0, 0, title, wb.add_format({"bold": True, "font_size": 13}))
    top = 2
    head_rows = 2 if cf else 1
    for j, f in enumerate(rf):
        if cf:
            ws.merge_range(top, j, top + 1, j, f, F.header())
        else:
            ws.write(top, j, f, F.header())
    c = nl
    for cap, span in G["header_top"]:
        span = span if cf else 1
        if span > 1:
            ws.merge_range(top, c, top, c + span - 1, cap, F.header())
        else:
            ws.write(top, c, cap, F.header())
        c += span
    for k, (text, hx) in enumerate(G["header_sub"]):
        ws.write(top + 1, nl + k, text, F.header(hx))
    r = top + head_rows
    for kind, labels, vals in G["lines"]:
        bold = kind != "row"
        fill = {"subtotal": SUB, "total": "D9D9D9"}.get(kind)
        for j, (text, hx) in enumerate(labels):
            if kind == "total":
                ws.write(r, j, text, F.header())
            elif kind == "subtotal" and j > 0:
                ws.write_blank(r, j, None, F.get(bg_color="#" + SUB))
            else:
                ws.write(r, j, text, F.label(hx, bold=bold))
        for k, (v, numfmt) in enumerate(vals):
            fmt = F.get(num_format=numfmt, bold=bold, **({"bg_color": "#" + fill} if fill else {}))
            if v is None:
                ws.write_blank(r, nl + k, None, fmt)
            else:
                ws.write_number(r, nl + k, v, fmt)
        r += 1
    for j, f in enumerate(rf):
        w = max([len(str(f))] + [len(str(l[1][j][0])) for l in G["lines"]] + [11])
        ws.set_column(j, j, min(45, w + 2))
    nvals = len(G["lines"][0][2]) if G["lines"] else 1
    ws.set_column(nl, nl + nvals - 1, 14)
    ws.freeze_panes(top + head_rows, nl)
    return G


def format_value(v, numfmt):
    """Number as the preview shows it (same formats as the sheet)."""
    if v is None:
        return ""
    if numfmt == "0.0%":
        return f"{v * 100:,.1f}%"
    if numfmt == "#,##0":
        return f"{v:,.0f}"
    return f"{v:,.2f}"


# ---------------------------------------------------------------------------
# workbook
# ---------------------------------------------------------------------------
def _write_source(wb, F, fields, rows, colors):
    ws = wb.add_worksheet("Source")
    data = [[r.get(f) for f in fields] for r in rows]
    ws.add_table(0, 0, max(1, len(rows)), len(fields) - 1, {
        "name": SOURCE_TABLE, "style": "Table Style Light 1",
        "columns": [{"header": f} for f in fields], "data": data or [[None] * len(fields)]})
    ws.freeze_panes(1, 0)
    for j, f in enumerate(fields):
        w = max([len(f)] + [len(str(r.get(f))) for r in rows[:300] if r.get(f) is not None])
        ws.set_column(j, j, max(8, min(40, w + 2)))
        scheme = colors.get(f)
        if scheme and rows:
            for value, hx in scheme.items():
                if value == BLANK:
                    continue
                crit = value if isinstance(value, (int, float)) else '"' + str(value).replace('"', '""') + '"'
                ws.conditional_format(1, j, len(rows), j, {
                    "type": "cell", "criteria": "==", "value": crit,
                    "format": wb.add_format({"bg_color": "#" + hx, "font_color": "#FFFFFF" if is_dark(hx) else "#000000"})})


def _write_legend(wb, F, rows, colors, orders, note):
    ws = wb.add_worksheet("Colours")
    ws.write(0, 0, "COLOUR CODING (from the layer symbology)", wb.add_format({"bold": True, "font_size": 13}))
    ws.write(1, 0, note)
    r = 3
    for f, scheme in colors.items():
        cnt = Counter(key(x.get(f)) for x in rows)
        for j, h in enumerate([f, "Colour", "HEX", "Features"]):
            ws.write(r, j, h, F.header())
        r += 1
        for v in sort_values(list(scheme), orders.get(f)):
            hx = scheme[v]
            ws.write(r, 0, v, F.label(hx))
            ws.write_blank(r, 1, None, F.get(bg_color="#" + hx))
            ws.write(r, 2, "#" + hx, F.get())
            ws.write(r, 3, cnt.get(v, 0), F.get(num_format="#,##0"))
            r += 1
        r += 1
    ws.set_column(0, 0, 34)
    ws.set_column(1, 3, 11)


def write_workbook(path, fields, rows, spec, colors, orders=None, native="auto",
                   title="PIVOT", note="", log=print):
    """native: 'auto' (Excel pivot if possible, else static), 'static', 'excel' (required).
    Returns 'excel' or 'static' (what was written)."""
    orders = orders or {}
    colors = {f: {v: h for v, h in s.items() if v != BLANK} for f, s in colors.items()}  # blanks stay plain
    if native in ("auto", "excel"):
        _build(path, fields, rows, spec, colors, orders, title, note, static=False)
        try:
            add_native_pivot(path, fields, spec, colors, title, log, orders)
            return "excel"
        except Exception as e:
            if native == "excel":
                raise
            log(f"Excel pivot table not available ({type(e).__name__}: {e}). Writing a static pivot instead.")
    _build(path, fields, rows, spec, colors, orders, title, note, static=True)
    return "static"


def _build(path, fields, rows, spec, colors, orders, title, note, static):
    if os.path.exists(path):
        os.remove(path)
    wb = xlsxwriter.Workbook(path, {"nan_inf_to_errors": True, "strings_to_urls": False,
                                    "strings_to_formulas": False})
    F = Formats(wb)
    piv = wb.add_worksheet("Pivot")
    if static:
        write_static_pivot(wb, piv, F, rows, spec, colors, orders, title + "  (static values)")
    _write_source(wb, F, fields, rows, colors)
    _write_legend(wb, F, rows, colors, orders, note)
    piv.activate()
    wb.close()


# ---------------------------------------------------------------------------
# native Excel pivot (Windows + Microsoft Excel, via COM)
# ---------------------------------------------------------------------------
XL_FUNC = {"sum": -4157, "mean": -4106, "min": -4139, "max": -4136, "count": -4112}
XL_PCT_OF_TOTAL = 8


def _bgr(hx):
    return int(hx[0:2], 16) + int(hx[2:4], 16) * 256 + int(hx[4:6], 16) * 65536


def add_native_pivot(path, fields, spec, colors, title, log=print, orders=None):
    import comtypes
    import comtypes.client
    import gc
    try:
        comtypes.CoInitialize()
    except OSError:
        pass
    xl = comtypes.client.CreateObject("Excel.Application", dynamic=True)
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.ScreenUpdating = False
        _native(xl, path, spec, colors, title, orders or {})
        log("Excel pivot table built (Data > Refresh All after editing Source).")
    finally:
        gc.collect()
        xl.Quit()
        del xl
        gc.collect()


def _native(xl, path, spec, colors, title, orders):
    wb = None
    try:
        dec, th = xl.International(3), xl.International(4)  # local number notation
        fmt_int, fmt_2, fmt_pct = f"#{th}##0", f"#{th}##0{dec}00", f"0{dec}0%"
        wb = xl.Workbooks.Open(os.path.abspath(path))
        ws = wb.Worksheets("Pivot")
        ws.Range("A1").Value2 = title
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 13
        pc = wb.PivotCaches().Create(1, SOURCE_TABLE, 5)
        pt = pc.CreatePivotTable(ws.Range("A3"), "ptVector")
        pt.RowAxisLayout(1)  # tabular
        try:
            pt.RepeatAllLabels(2)  # xlRepeatLabels (Excel 2010+)
        except Exception:
            pass
        pt.TableStyle2 = "PivotStyleLight16"
        rf, cf = spec["row_fields"], spec.get("col_field")
        for i, f in enumerate(rf):
            pf = pt.PivotFields(f)
            pf.Orientation = 1
            pf.Position = i + 1
            if i > 0:
                pf.Subtotals[1] = False  # indexed COM property; 1 = Automatic -> no subtotals
        if cf:
            pt.PivotFields(cf).Orientation = 2
        n = 0
        for f, agg in spec["measures"]:
            df = pt.AddDataField(pt.PivotFields(f), caption(f, agg), XL_FUNC[agg])
            df.NumberFormat = fmt_int if number_format(f, agg) == "#,##0" else fmt_2
            n += 1
            if spec.get("percent") and agg in ADDITIVE:
                df = pt.AddDataField(pt.PivotFields(f), f"% of total ({caption(f, agg)})", XL_FUNC[agg])
                df.Calculation = XL_PCT_OF_TOTAL
                df.NumberFormat = fmt_pct
                n += 1
        if n > 1 and cf:
            pt.DataPivotField.Orientation = 2  # Values on columns
        for f in rf + ([cf] if cf else []):  # legend order (pivot items are text)
            pos = 1
            for v in orders.get(f) or []:
                try:
                    pt.PivotFields(f).PivotItems(str(v)).Position = pos
                    pos += 1
                except Exception:
                    pass
        # colour coding: conditional formats survive refresh / re-sorting
        for f in rf:
            _cf(ws, ws.Columns(pt.PivotFields(f).DataRange.Column), colors.get(f))
            if f == rf[0] and len(rf) > 1:  # "<value> Total" subtotal labels
                _cf(ws, ws.Columns(pt.PivotFields(f).DataRange.Column), colors.get(f), suffix=" Total")
        if cf and colors.get(cf):
            tr = pt.TableRange1
            hdr_row = pt.PivotFields(cf).DataRange.Row
            _cf(ws, ws.Range(ws.Cells(hdr_row, tr.Column), ws.Cells(hdr_row, tr.Column + tr.Columns.Count + 50)),
                colors.get(cf))
        ws.Range("A3:ZZ5000").Columns.AutoFit()
        ws.Activate()
        wb.Save()
    finally:
        if wb is not None:
            wb.Close(False)


def _cf(ws, rng, scheme, suffix=""):
    for value, hx in (scheme or {}).items():
        if value == BLANK:
            continue
        if isinstance(value, (int, float)) and not suffix:
            formula = "=" + repr(value)
        else:
            formula = '="' + (str(value) + suffix).replace('"', '""') + '"'
        fc = rng.FormatConditions.Add(1, 3, formula)  # xlCellValue, xlEqual
        fc.Interior.Color = _bgr(hx)
        fc.Font.Color = _bgr("FFFFFF" if is_dark(hx) else "000000")
        fc.StopIfTrue = False
