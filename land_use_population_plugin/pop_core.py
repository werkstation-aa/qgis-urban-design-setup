"""Land-use population calculation - no QGIS imports, so it can be tested on its own.

Reference values live in pop_reference.json (published standards only: SNI 03-1733-2004,
Permen PU 26/PRT/M/2008, Kepmen Kimpraswil 403/2002, Permenpera 11/2008 and the HCA
Employment Density Guide 2015).
"""

import copy
import datetime
import json
import math
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
LIBRARY_PATH = os.path.join(HERE, "pop_reference.json")

# method -> [(param key, label, unit, decimals)]
METHOD_PARAMS = {
    "none": [],
    "density": [("coef", "Density", "jiwa/ha", 1)],
    "units_lot": [("lot_m2", "Lot size", "m2/unit", 1), ("net_ratio", "Saleable ratio", "0-1", 2),
                  ("pph", "Occupancy", "jiwa/unit", 2)],
    "units_gfa": [("unit_m2", "Unit size", "m2 GFA/unit", 1), ("eff", "Net efficiency", "0-1", 2),
                  ("pph", "Occupancy", "jiwa/unit", 2)],
    "gfa_per_person": [("coef", "Floor area per person", "m2 GFA/jiwa", 3)],
    "rate_per_gfa": [("coef", "Persons per m2 GFA", "jiwa/m2", 4)],
    "rate_per_ha": [("coef", "Persons per ha", "jiwa/ha", 2)],
    "hotel": [("room_m2", "Room size", "m2/key", 1), ("eff", "Net efficiency", "0-1", 2),
              ("occ", "Occupancy", "0-1", 2), ("guests", "Guests per room", "jiwa", 2),
              ("staff_per_key", "Staff per key", "jiwa/key", 2)],
    "fixed": [("coef", "Persons per parcel", "jiwa", 0)],
}
NEEDS_GFA = {"units_gfa", "gfa_per_person", "rate_per_gfa", "hotel"}
METHOD_LABEL = {
    "none": "No population", "density": "Area x density", "units_lot": "Lots x occupancy",
    "units_gfa": "Units x occupancy", "gfa_per_person": "GFA / m2 per person",
    "rate_per_gfa": "GFA x rate", "rate_per_ha": "Area x rate", "hotel": "Hotel keys", "fixed": "Fixed",
}
STATUS_COLORS = {  # fill colours (hex) used in the window and the workbook
    "verified": "C6EFCE", "derived": "DDEBF7", "input": "FFF2CC",
}
RESIDENT = "Resident"


def load_library(path=None):
    with open(path or LIBRARY_PATH, encoding="utf-8") as fh:
        lib = json.load(fh)
    lib["by_id"] = {p["id"]: p for p in lib["profiles"]}
    return lib


def _kw_hit(kw, text):
    kw = kw.lower().strip()
    if len(kw) <= 4:  # short codes (sd, smp, rth, kpi) must be whole words
        return re.search(r"(?<![a-z0-9])" + re.escape(kw) + r"(?![a-z0-9])", text) is not None
    return kw in text


def auto_match(value, lib):
    """Profile id for a land-use class name: the longest matching keyword wins."""
    text = str(value or "").lower()
    best, best_len = "NONE" if not text else None, 0
    for p in lib["profiles"]:
        for kw in p.get("match", []):
            if len(kw) > best_len and _kw_hit(kw, text):
                best, best_len = p["id"], len(kw)
    return best or "NONE"


def default_assignment(value, lib):
    pid = auto_match(value, lib)
    return {"profile": pid, "params": copy.deepcopy(lib["by_id"][pid]["params"]), "auto": True}


def formula_text(method, params):
    p = params or {}
    f = lambda k: fmt_num(p.get(k, 0))
    return {
        "none": "0",
        "density": f"Area(ha) x {f('coef')} jiwa/ha",
        "units_lot": f"Area x {f('net_ratio')} / {f('lot_m2')} m2 x {f('pph')} jiwa/unit",
        "units_gfa": f"GFA x {f('eff')} / {f('unit_m2')} m2 x {f('pph')} jiwa/unit",
        "gfa_per_person": f"GFA / {f('coef')} m2/jiwa",
        "rate_per_gfa": f"GFA x {f('coef')} jiwa/m2",
        "rate_per_ha": f"Area(ha) x {f('coef')} jiwa/ha",
        "hotel": (f"Keys = GFA x {f('eff')} / {f('room_m2')} m2; guests = keys x {f('occ')} x {f('guests')}; "
                  f"staff = keys x {f('staff_per_key')}"),
        "fixed": f"{f('coef')} jiwa per parcel",
    }.get(method, "?")


def fmt_num(v):
    if v is None:
        return "-"
    if abs(v - round(v)) < 1e-9:
        return f"{int(round(v)):,}"
    return f"{v:,.4f}".rstrip("0").rstrip(".")


def compute(method, params, pop_type, area_m2, gfa_m2):
    """Population of one parcel -> dict(resi, visit, units, ok, msg)."""
    p = params or {}
    g = lambda k: float(p.get(k) or 0)
    area_m2 = float(area_m2 or 0)
    out = dict(resi=0.0, visit=0.0, units=None, ok=True, msg="")
    if method in NEEDS_GFA and not gfa_m2:
        out.update(ok=False, msg="No GFA")
        return out
    gfa = float(gfa_m2 or 0)
    pop = 0.0
    if method == "density":
        pop = area_m2 / 10000.0 * g("coef")
    elif method == "rate_per_ha":
        pop = area_m2 / 10000.0 * g("coef")
    elif method == "units_lot":
        units = area_m2 * g("net_ratio") / g("lot_m2") if g("lot_m2") else 0
        out["units"], pop = units, units * g("pph")
    elif method == "units_gfa":
        units = gfa * g("eff") / g("unit_m2") if g("unit_m2") else 0
        out["units"], pop = units, units * g("pph")
    elif method == "gfa_per_person":
        pop = gfa / g("coef") if g("coef") else 0
    elif method == "rate_per_gfa":
        pop = gfa * g("coef")
    elif method == "fixed":
        pop = g("coef")
    elif method == "hotel":
        keys = gfa * g("eff") / g("room_m2") if g("room_m2") else 0
        out["units"] = keys
        pop = keys * g("occ") * g("guests") + keys * g("staff_per_key")
    if pop_type == RESIDENT:
        out["resi"] = pop
    elif method != "none":
        out["visit"] = pop
    return out


def resolve_gfa(gfa_value, far_value, area_m2):
    """GFA used for a parcel: the GFA field when > 0, else area x FAR. Returns (gfa, source, qa)."""
    gfa = _num(gfa_value)
    far = _num(far_value)
    qa = ""
    if gfa and far and area_m2:
        expect = area_m2 * far
        if expect and abs(gfa - expect) / expect > 0.10:
            qa = f"GFA field {gfa:,.0f} differs from Area x FAR {expect:,.0f} by {100 * (gfa - expect) / expect:+.0f}%"
    if gfa:
        return gfa, "GFA field", qa
    if far and area_m2:
        return area_m2 * far, "Area x FAR", qa
    return None, "", qa


def _num(v):
    try:
        f = float(v)
        return None if math.isnan(f) or f <= 0 else f
    except (TypeError, ValueError):
        return None


def summarise(records, assignments, lib):
    """records: [{cls, area_m2, gfa}] -> per-class summary rows (ordered by class name)."""
    rows = {}
    for r in records:
        cls = r["cls"]
        a = assignments.get(cls) or default_assignment(cls, lib)
        prof = lib["by_id"].get(a["profile"], lib["by_id"]["NONE"])
        res = compute(prof["method"], a["params"], prof["pop_type"], r["area_m2"], r["gfa"])
        s = rows.setdefault(cls, dict(cls=cls, n=0, area_m2=0.0, gfa=0.0, resi=0.0, visit=0.0,
                                      units=None, no_gfa=0, profile=prof, params=a["params"]))
        s["n"] += 1
        s["area_m2"] += r["area_m2"] or 0
        s["gfa"] += r["gfa"] or 0
        s["resi"] += res["resi"]
        s["visit"] += res["visit"]
        if res["units"] is not None:
            s["units"] = (s["units"] or 0) + res["units"]
        if not res["ok"]:
            s["no_gfa"] += 1
    return [rows[k] for k in sorted(rows, key=lambda x: str(x))]


# --------------------------------------------------------------------------- compact assumption table
ASSUMPTION_FOR = {"Resident": "Resident"}  # everything else -> "Visitor / Worker"
COMPACT_COLS = ["Assumption for", "Land Use", "Program", "Coefficient", "Units", "Source"]
FOR_COLORS = {"Resident": "FFC000", "Visitor / Worker": "B4C7E7"}
FALLBACK_COLORS = ["4E79A7", "F28E2B", "76B7B2", "EDC948", "B07AA1", "FF9DA7", "9C755F", "BAB0AC",
                   "59A14F", "E15759"]


def fmt_id(v, decimals=3):
    """Number in Indonesian style as in the example slide: 3,5 / 1104 / 16,66."""
    if v is None:
        return "-"
    s = f"{float(v):.{decimals}f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


def pct(v):
    return f"{fmt_id(float(v) * 100, 1)}%"


def source_text(lib, prof, full=False):
    """Citation of a profile: document (+ table / clause), then 'cf.' comparison document.
    full=True gives the complete bibliographic entries, else the short document names."""
    names = lib["sources"] if full else lib.get("sources_short", lib["sources"])
    key = prof.get("source", "-")
    out = names.get(key, key)
    if prof.get("ref"):
        out += f", {prof['ref']}" if not full else f"  [{prof['ref']}]"
    if prof.get("cf"):
        cf = names.get(prof["cf"], prof["cf"])
        out += f"; cf. {cf}" + (f", {prof['cf_ref']}" if prof.get("cf_ref") and not full else
                                f"  [{prof['cf_ref']}]" if prof.get("cf_ref") else "")
    return out


def program_name(prof):
    return prof.get("program") or prof["name"].split(" - ")[0].strip()


def compact_coef(method, params):
    """(coefficient, units, detail for the Program cell) for the compact table."""
    p = params or {}
    if method in ("density", "rate_per_ha"):
        return fmt_id(p.get("coef")), "People/Ha", ""
    if method == "gfa_per_person":
        return fmt_id(p.get("coef")), "m²/person", ""
    if method == "rate_per_gfa":
        return fmt_id(p.get("coef"), 4), "People/m² GFA", ""
    if method == "units_lot":
        return (fmt_id(p.get("pph")), "People/unit",
                f"lot {fmt_id(p.get('lot_m2'))} m², {pct(p.get('net_ratio', 0))} saleable")
    if method == "units_gfa":
        return (fmt_id(p.get("pph")), "People/unit",
                f"{fmt_id(p.get('unit_m2'))} m² GFA/unit, {pct(p.get('eff', 0))} net")
    if method == "hotel":
        return (fmt_id(p.get("room_m2")), "m²/room",
                f"{pct(p.get('occ', 0))} occ., {fmt_id(p.get('guests'))} guests/room, "
                f"{fmt_id(p.get('staff_per_key'))} staff/key")
    if method == "fixed":
        return fmt_id(p.get("coef")), "People/parcel", ""
    return "-", "-", ""


def compact_rows(summary, lib, colors=None):
    """One row per land-use class with population, grouped Resident first, then Visitor / Worker.
    -> [dict(assumption_for, land_use, program, coef, units, source, status, color, for_color)]"""
    colors = colors or {}
    out = []
    fb = iter(FALLBACK_COLORS * 20)
    for s in summary:
        prof = s["profile"]
        if prof["method"] == "none":
            continue
        coef, units, detail = compact_coef(prof["method"], s["params"])
        changed = any(abs(float(s["params"].get(k) or 0) - float(v or 0)) > 1e-9 for k, v in prof["params"].items())
        src = source_text(lib, prof)
        afor = ASSUMPTION_FOR.get(prof["pop_type"], "Visitor / Worker")
        prog = program_name(prof) + (f" ({detail})" if detail else "")
        out.append(dict(assumption_for=afor, land_use=s["cls"], program=prog, coef=coef, units=units,
                        source=src + (" (adjusted)" if changed else ""), status=prof["status"],
                        color=(colors.get(s["cls"]) or next(fb)).upper().lstrip("#"),
                        for_color=FOR_COLORS[afor], resi=s["resi"], visit=s["visit"]))
    order = {"Resident": 0, "Visitor / Worker": 1}
    out.sort(key=lambda r: order[r["assumption_for"]])  # stable: keeps the class order inside a group
    return out


def spans(rows, key):
    """[(start, end)] runs of equal rows[key] (inclusive), for merged cells."""
    res, start = [], 0
    for i in range(1, len(rows) + 1):
        if i == len(rows) or rows[i][key] != rows[start][key]:
            res.append((start, i - 1))
            start = i
    return res


def is_dark(hex_):
    h = hex_.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return (0.299 * r + 0.587 * g + 0.114 * b) < 150


def sni_needs(population, lib, gross_density=None, apply_reduction=False):
    """Facility need per SNI 03-1733-2004 for a resident population."""
    red = 0.0
    if apply_reduction and gross_density:
        red = 0.30 if gross_density > 400 else 0.15 if gross_density > 200 else 0.0
    out = []
    for grp, name, level, support, floor, land, radius, status, note in lib["sni_facilities"]["rows"]:
        exact = population / support if support else 0
        # round up; under half a catchment the facility is assumed shared with the surrounding area
        units = math.ceil(exact - 1e-9) if exact >= 0.5 else 0
        out.append(dict(group=grp, name=name, level=level, support=support, floor=floor, land=land,
                        radius=radius, status=status, note=note, exact=exact, units=units,
                        floor_total=(floor or 0) * units if floor else None,
                        land_total=(land or 0) * units * (1 - red) if land else None, reduction=red))
    return out


# --------------------------------------------------------------------------- Excel
def write_workbook(path, fields, rows, summary, lib, meta, sni=None, compact=None):
    """Sheets 'Land Use' (attribute table with population fields), 'Assumption Summary' (compact
    table as on the slide), 'Population Assumptions' (full detail) and optional 'SNI Facility Needs'.
    Needs openpyxl (bundled with QGIS 3.x on Windows)."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="305496")
    hdr_font = Font(bold=True, color="FFFFFF")
    bold = Font(bold=True)
    wrap = Alignment(wrap_text=True, vertical="top")

    def header(ws, r, labels, c0=1):
        for i, lab in enumerate(labels):
            c = ws.cell(r, c0 + i, lab)
            c.fill, c.font, c.border = hdr_fill, hdr_font, box
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

    def widths(ws, ws_widths):
        for i, w in enumerate(ws_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    wb = Workbook()
    # ---- Land Use
    ws = wb.active
    ws.title = "Land Use"
    header(ws, 1, fields)
    pop_cols = {f for f in fields if f.startswith("Pop_")}
    for r_i, row in enumerate(rows, 2):
        for c_i, f in enumerate(fields, 1):
            v = row.get(f)
            c = ws.cell(r_i, c_i, v)
            c.border = box
            if isinstance(v, float):
                c.number_format = "#,##0.00" if abs(v) < 1000 else "#,##0"
            elif isinstance(v, int):
                c.number_format = "#,##0"
            if f in pop_cols:
                c.fill = PatternFill("solid", fgColor="FFF2CC")
    last = len(rows) + 1
    tot_r = last + 1
    ws.cell(tot_r, 1, "TOTAL").font = bold
    for c_i, f in enumerate(fields, 1):
        if f in ("Pop_Resi", "Pop_Visit", "Pop_Total", "Pop_Units", "Pop_GFA") or f.lower() in (
                meta.get("area_field", "").lower(), "area_sqm", "area_ha", "gfa_sqm"):
            col = get_column_letter(c_i)
            c = ws.cell(tot_r, c_i, f"=SUBTOTAL(109,{col}2:{col}{last})")
            c.font, c.number_format = bold, "#,##0"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(fields))}{last}"
    for c_i, f in enumerate(fields, 1):
        ws.column_dimensions[get_column_letter(c_i)].width = max(10, min(40, len(f) + 4))

    # ---- Assumption Summary (compact, same layout as the PowerPoint table)
    if compact is None:
        compact = compact_rows(summary, lib)
    wc = wb.create_sheet("Assumption Summary")
    wc.sheet_view.showGridLines = False
    wc["A1"] = "Population Assumption"
    wc["A1"].font = Font(bold=True, size=16)
    black = PatternFill("solid", fgColor="000000")
    line = Side(style="thin", color="D9D9D9")
    for j, lab in enumerate(COMPACT_COLS, 1):
        c = wc.cell(3, j, lab)
        c.fill, c.font = black, Font(bold=True, color="FFFFFF", size=11)
        c.alignment = Alignment(vertical="center")
    keys = ["assumption_for", "land_use", "program", "coef", "units", "source"]
    for i, r in enumerate(compact, 4):
        for j, k in enumerate(keys, 1):
            c = wc.cell(i, j, r[k])
            c.font = Font(size=11, bold=j <= 2)
            c.alignment = Alignment(vertical="center", wrap_text=j >= 5, horizontal="right" if j == 4 else None)
            if j >= 3:
                c.border = Border(top=line, bottom=line)
        st = wc.cell(i, 7, r["status"])
        st.fill = PatternFill("solid", fgColor=STATUS_COLORS.get(r["status"], "FFFFFF"))
        st.font, st.alignment = Font(size=9, italic=True), Alignment(vertical="center", horizontal="center")
    if compact:
        wc.cell(3, 7, "Status").font = Font(bold=True, size=9, color="808080")
    for key, col in (("assumption_for", 1), ("land_use", 2)):
        for a, b in spans(compact, key):
            r = compact[a]
            fill_hex = r["for_color"] if key == "assumption_for" else r["color"]
            for i in range(a, b + 1):
                wc.cell(4 + i, col).fill = PatternFill("solid", fgColor=fill_hex)
            top = wc.cell(4 + a, col)
            top.font = Font(size=11, bold=True, color="FFFFFF" if is_dark(fill_hex) else "000000")
            top.alignment = Alignment(vertical="center", wrap_text=True)
            if b > a:
                wc.merge_cells(start_row=4 + a, start_column=col, end_row=4 + b, end_column=col)
    end = 4 + len(compact)
    res = sum(round(r["resi"]) for r in compact)
    vis = sum(round(r["visit"]) for r in compact)
    wc.cell(end + 1, 1, f"Residents {res:,} · Visitors / workers {vis:,} · Total {res + vis:,}").font = Font(
        italic=True, size=10)
    wc.cell(end + 2, 1, "Full formulas, sources and notes: sheet 'Population Assumptions'.").font = Font(
        italic=True, size=10, color="808080")
    widths(wc, [18, 20, 30, 12, 34, 44, 11])
    for i in range(4, end):
        wc.row_dimensions[i].height = 20

    # ---- Population Assumptions
    wa = wb.create_sheet("Population Assumptions")
    wa["A1"] = "Population assumptions used in this run"
    wa["A1"].font = Font(bold=True, size=14)
    info = [("Layer", meta.get("layer")), ("Land-use field", meta.get("class_field")),
            ("Area source", meta.get("area_source")), ("GFA source", meta.get("gfa_source")),
            ("Features", meta.get("features")), ("Generated", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")),
            ("Reference", lib.get("built_from"))]
    for i, (k, v) in enumerate(info, 2):
        wa.cell(i, 1, k).font = bold
        wa.cell(i, 2, v)
    r0 = len(info) + 3
    cols = ["Land-use class", "Parcels", "Area (ha)", "GFA (m2)", "Assumption", "Group", "Method",
            "Formula", "Parameters", "Population type", "Units / keys", "Residents", "Visitors / workers",
            "Total population", "Status", "Source", "Note"]
    header(wa, r0, cols)
    t_res = t_vis = t_area = t_gfa = 0
    for i, s in enumerate(summary, r0 + 1):
        prof = s["profile"]
        par = "; ".join(f"{lab} = {fmt_num(s['params'].get(k))} {u}" for k, lab, u, _ in
                        METHOD_PARAMS.get(prof["method"], []))
        vals = [s["cls"], s["n"], s["area_m2"] / 10000, s["gfa"], prof["name"], prof["group"],
                METHOD_LABEL.get(prof["method"], prof["method"]), formula_text(prof["method"], s["params"]),
                par, prof["pop_type"], round(s["units"]) if s["units"] is not None else None,
                round(s["resi"]), round(s["visit"]), round(s["resi"]) + round(s["visit"]), prof["status"],
                source_text(lib, prof, full=True),
                prof["note"] + (f"  [{s['no_gfa']} parcel(s) without GFA -> 0]" if s["no_gfa"] else "")]
        for j, v in enumerate(vals, 1):
            c = wa.cell(i, j, v)
            c.border, c.alignment = box, wrap
            if j in (3,):
                c.number_format = "#,##0.00"
            elif j in (2, 4, 11, 12, 13, 14):
                c.number_format = "#,##0"
        wa.cell(i, 15).fill = PatternFill("solid", fgColor=STATUS_COLORS.get(prof["status"], "FFFFFF"))
        t_res += round(s["resi"])
        t_vis += round(s["visit"])
        t_area += s["area_m2"]
        t_gfa += s["gfa"]
    tr = r0 + len(summary) + 1
    for j, v in ((1, "TOTAL"), (3, t_area / 10000), (4, t_gfa), (12, t_res), (13, t_vis), (14, t_res + t_vis)):
        c = wa.cell(tr, j, v)
        c.font, c.border = bold, box
        c.number_format = "#,##0.00" if j == 3 else "#,##0"
    if t_area:
        wa.cell(tr + 1, 1, "Gross resident density (jiwa/ha of all parcels)").font = bold
        wa.cell(tr + 1, 12, t_res / (t_area / 10000)).number_format = "#,##0.0"
    r = tr + 3
    for title, d in (("Methods", {k: v for k, v in lib["methods"].items()}),
                     ("Status legend", lib["status_legend"]), ("Sources", lib["sources"])):
        wa.cell(r, 1, title).font = Font(bold=True, size=12)
        r += 1
        for k, v in d.items():
            wa.cell(r, 1, k).font = bold
            wa.cell(r, 2, v)
            if title == "Status legend":
                wa.cell(r, 1).fill = PatternFill("solid", fgColor=STATUS_COLORS.get(k, "FFFFFF"))
            r += 1
        r += 1
    widths(wa, [26, 9, 11, 13, 34, 16, 18, 40, 44, 15, 11, 12, 14, 14, 12, 40, 60])
    wa.freeze_panes = wa.cell(r0 + 1, 2)

    # ---- optional SNI facility needs
    if sni is not None:
        wn = wb.create_sheet("SNI Facility Needs")
        wn["A1"] = f"SNI 03-1733-2004 facility needs for {meta.get('residents', 0):,.0f} residents"
        wn["A1"].font = Font(bold=True, size=14)
        wn["A2"] = lib["sni_facilities"]["source"] + ". " + lib["sni_facilities"]["density_reduction"]
        cols = ["Group", "Facility", "Level", "Population per unit", "Min floor (m2)", "Min land (m2)",
                "Radius (m)", "Units (exact)", "Units (planned)", "Floor need (m2)", "Land need (m2)",
                "Land need (ha)", "Status", "Note"]
        header(wn, 4, cols)
        for i, n in enumerate(sni, 5):
            vals = [n["group"], n["name"], n["level"], n["support"], n["floor"], n["land"], n["radius"],
                    round(n["exact"], 2), n["units"], n["floor_total"], n["land_total"],
                    (n["land_total"] / 10000) if n["land_total"] else None, n["status"], n["note"]]
            for j, v in enumerate(vals, 1):
                c = wn.cell(i, j, v)
                c.border = box
                if j in (4, 5, 6, 7, 10, 11):
                    c.number_format = "#,##0"
                elif j == 12:
                    c.number_format = "#,##0.00"
            wn.cell(i, 13).fill = PatternFill("solid", fgColor=STATUS_COLORS.get(n["status"], "FFFFFF"))
        widths(wn, [28, 42, 12, 12, 11, 11, 10, 11, 11, 13, 13, 11, 11, 60])
        wn.freeze_panes = "C5"

    wb.save(path)
    return path
