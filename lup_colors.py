"""
Colour scheme for land-use categories, shared by the Excel conditional formatting,
the slides and the QGIS layer symbology.

For each value of the chosen "colour field" (e.g. SUBZONA or ZONA) the colour comes
from the first source that has one, in the chosen priority order:

    manual override  (set in the Export Land Use window; always wins)
    layer symbology  (categorised renderer of the QGIS layer on that field)
    colour library   (RDTR / ATR-BPN colour code workbook, matched by name, code,
                      alias or close spelling)
    project palette  (colours of the sample deck, lup_pipeline.COLORS)

Settings (library path, field, priority, aliases, overrides) live in
lup_colours.json next to this file.
"""

import difflib
import json
import os
import re
from collections import OrderedDict, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_NAME = "lup_colours.json"
LIBRARY_NAME = "RDTR_Colour Code.xlsx"

PRIORITIES = OrderedDict([
    ("layer", ("Layer symbology → colour library → project palette", ["layer", "library", "palette"])),
    ("library", ("Colour library → layer symbology → project palette", ["library", "layer", "palette"])),
    ("palette", ("Project palette → colour library", ["palette", "library", "layer"])),
])

# English / project terms -> RDTR code (Kode SZ or Kode Z). Editable per value in the window.
DEFAULT_ALIASES = {
    "industrial": "KPI", "industry": "KPI", "kawasan industri": "KPI",
    "commercial": "K-1", "retail": "K-1", "trade": "K-1",
    "mice": "K-2", "hotel": "W", "hospitality": "W", "tourism": "W", "resort": "W",
    "office": "KT", "residential": "R-3", "housing": "R-3", "resi": "R-3",
    "dormitory": "R-2", "dorm": "R-2",
    "golf": "RTH-2", "park": "RTH-2", "green": "RTH-8", "green belt": "RTH-8", "buffer": "RTH-8",
    "cemetery": "RTH-7", "open space": "RTNH", "plaza": "RTNH",
    "road": "BJ", "street": "BJ", "port": "TR", "transport": "TR", "transportation": "TR",
    "terminal": "TR", "utility": "PL-3", "utilities": "PL-3", "wtp": "PL-3", "stp": "PL-4",
    "warehouse": "PL-6", "logistics": "PL-6", "public amenity": "SPU-1", "public facility": "SPU-1",
    "mixed use": "C-1", "water": "BA", "river": "BA", "lake": "BA", "mangrove": "EM",
    "agriculture": "P-1", "plantation": "P-3", "forest": "HP", "conservation": "CA",
    "defense": "HK", "military": "HK", "waste": "PP", "power plant": "PTL",
    # Indonesian zone names used without "Zona" / with short forms
    "lainnya": "PL", "rth": "RTH", "ruang terbuka hijau": "RTH-2",
}


def norm(s):
    s = str(s or "").strip().lower().replace("&", " dan ")
    s = re.sub(r"\s+", " ", s)
    return s[5:] if s.startswith("zona ") else s


def is_dark(hex_):
    r, g, b = (int(hex_[i:i + 2], 16) for i in (0, 2, 4))
    return (0.299 * r + 0.587 * g + 0.114 * b) < 150


def rgb_text_to_hex(t):
    m = re.findall(r"\d+", str(t or ""))
    if len(m) < 3:
        return None
    return "".join(f"{max(0, min(255, int(v))):02X}" for v in m[:3])


def hex_to_rgb_text(h):
    return ",".join(str(int(h[i:i + 2], 16)) for i in (0, 2, 4))


# ---------------------------------------------------------------------------
# config
# ---------------------------------------------------------------------------
def config_path(lib_dir=None):
    return os.path.join(lib_dir or HERE, CONFIG_NAME)


def load_config(lib_dir=None):
    cfg = {"library": "", "field": "SUBZONA", "priority": "layer", "aliases": {}, "overrides": {}}
    try:
        with open(config_path(lib_dir), encoding="utf-8") as f:
            cfg.update(json.load(f))
    except (OSError, ValueError):
        pass
    return cfg


def save_config(cfg, lib_dir=None):
    try:
        with open(config_path(lib_dir), "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except OSError:
        pass


def library_path(cfg, lib_dir=None):
    for p in (cfg.get("library"), os.path.join(lib_dir or HERE, LIBRARY_NAME)):
        if p and os.path.exists(p):
            return p
    return None


# ---------------------------------------------------------------------------
# colour library (RDTR colour code workbook)
# ---------------------------------------------------------------------------
_lib_cache = {}


def load_library(path):
    """Entries: dict(zona, kode_z, subzona, kode_sz, hex, label). Reads the sheet layout
    'Zona | Kode Z | Sub-Zona | Kode SZ | RGB' (headers may sit over merged columns)."""
    if not path or not os.path.exists(path):
        return []
    key = (path, os.path.getmtime(path))
    if key in _lib_cache:
        return _lib_cache[key]
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    entries = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        hdr_i, cols = None, {}
        for i, row in enumerate(rows[:30]):
            labels = [norm(v) for v in row]
            if "rgb" in labels:
                hdr_i = i
                for j, lab in enumerate(labels):
                    k = {"zona": "zona", "kode z": "kode_z", "sub-zona": "subzona", "subzona": "subzona",
                         "sub zona": "subzona", "kode sz": "kode_sz", "rgb": "rgb"}.get(lab)
                    if k and k not in cols:
                        cols[k] = j
                break
        if hdr_i is None:
            continue

        def get(row, k, spread=1):
            j = cols.get(k)
            if j is None:
                return None
            for jj in range(j, min(len(row), j + 1 + spread)):  # header over merged cells
                if row[jj] not in (None, ""):
                    return str(row[jj]).strip()
            return None

        zona = None
        for row in rows[hdr_i + 1:]:
            hx = rgb_text_to_hex(get(row, "rgb", 0))
            if not hx:
                continue
            zona = get(row, "zona") or zona
            e = dict(zona=zona or "", kode_z=get(row, "kode_z", 0) or "", subzona=get(row, "subzona") or "",
                     kode_sz=get(row, "kode_sz", 0) or "", hex=hx)
            e["label"] = f"{e['kode_sz'] or e['kode_z']}  {e['subzona'] or e['zona']}"
            entries.append(e)
        if entries:
            break
    _lib_cache[key] = entries
    return entries


def find_code(lib, code):
    c = norm(code)
    for e in lib:
        if norm(e["kode_sz"]) == c:
            return e
    for e in lib:  # zone code -> first sub-zone of that zone
        if norm(e["kode_z"]) == c:
            return e
    return None


def match_library(value, lib, aliases=None):
    """-> (entry, how) or (None, None)."""
    if not lib or value in (None, ""):
        return None, None
    v = norm(value)
    user = {norm(k): c for k, c in (aliases or {}).items()}
    if v in user:
        if user[v] in ("", None, "-"):
            return None, None
        e = find_code(lib, user[v])
        if e:
            return e, "chosen"
    for e in lib:
        if v == norm(e["subzona"]) or v == norm(e["kode_sz"]):
            return e, "name"
    for e in lib:
        if v == norm(e["zona"]) or v == norm(e["kode_z"]):
            return e, "zone"
    if v in DEFAULT_ALIASES:
        e = find_code(lib, DEFAULT_ALIASES[v])
        if e:
            return e, "alias"
    names = {}
    for e in lib:
        names.setdefault(norm(e["subzona"]), e)
        names.setdefault(norm(e["zona"]), e)
    close = difflib.get_close_matches(v, list(names), n=1, cutoff=0.82)
    if close:
        return names[close[0]], "similar"
    return None, None


# ---------------------------------------------------------------------------
# QGIS layer symbology (only used inside QGIS)
# ---------------------------------------------------------------------------
def renderer_field(layer):
    try:
        r = layer.renderer()
        if r is not None and r.type() == "categorizedSymbol":
            return r.classAttribute()
    except Exception:
        pass
    return None


def layer_colors(layer, field):
    """{value(str): HEX} from a categorised renderer on `field`, else {}."""
    try:
        r = layer.renderer()
        if r is None or r.type() != "categorizedSymbol" or r.classAttribute() != field:
            return {}
        out = {}
        for c in r.categories():
            if c.symbol() is not None and c.value() not in (None, ""):
                out[str(c.value())] = c.symbol().color().name()[1:].upper()
        return out
    except Exception:
        return {}


def apply_to_layer(layer, field, scheme):
    """Write `scheme` {value: HEX} to the layer as a categorised renderer, keeping the
    existing symbol style (outline etc.) when the layer is already categorised."""
    from qgis.core import QgsCategorizedSymbolRenderer, QgsRendererCategory, QgsSymbol
    from qgis.PyQt.QtGui import QColor
    old = layer.renderer()
    base = None
    # clone straight away: categories() returns temporary copies, and the old renderer
    # (which owns its symbols) is deleted by setRenderer() below
    if old is not None and old.type() == "categorizedSymbol":
        old_cats = old.categories()
        if old_cats and old_cats[0].symbol() is not None:
            base = old_cats[0].symbol().clone()
    elif old is not None and old.type() == "singleSymbol" and old.symbol() is not None:
        base = old.symbol().clone()
    if base is None:
        base = QgsSymbol.defaultSymbol(layer.geometryType())
    cats = []
    for value, hx in scheme.items():
        sym = base.clone()
        sym.setColor(QColor("#" + hx))
        cats.append(QgsRendererCategory(value, sym, str(value)))
    layer.setRenderer(QgsCategorizedSymbolRenderer(field, cats))
    layer.triggerRepaint()
    layer.emitStyleChanged()


# ---------------------------------------------------------------------------
# resolve
# ---------------------------------------------------------------------------
FALLBACK = ["4E79A7", "F28E2B", "76B7B2", "EDC948", "B07AA1", "FF9DA7", "9C755F", "BAB0AC", "59A14F", "E15759"]


def resolve(values, field, cfg, lib, layer_scheme=None, palette=None):
    """values -> OrderedDict value -> dict(hex, source, lib_entry, lib_how, layer_hex)."""
    order = PRIORITIES.get(cfg.get("priority"), PRIORITIES["layer"])[1]
    overrides = (cfg.get("overrides") or {}).get(field, {})
    layer_scheme = layer_scheme or {}
    palette = palette or {}
    out, fb = OrderedDict(), 0
    for v in values:
        key = str(v)
        e, how = match_library(v, lib, cfg.get("aliases"))
        cand = {"layer": layer_scheme.get(key), "palette": palette.get(key),
                "library": e["hex"] if e else None}
        if overrides.get(key):
            hx, src = overrides[key].upper(), "manual"
        else:
            hx, src = None, None
            for s in order:
                if cand[s]:
                    hx, src = cand[s].upper(), s
                    break
            if not hx:
                hx, src = FALLBACK[fb % len(FALLBACK)], "auto"
                fb += 1
        out[key] = dict(hex=hx, source=src, lib_entry=e, lib_how=how, layer_hex=layer_scheme.get(key))
    return out


def source_text(info):
    s = info["source"]
    if s == "library" and info["lib_entry"]:
        e = info["lib_entry"]
        return f"Library {e['kode_sz'] or e['kode_z']} ({info['lib_how']})"
    return {"layer": "Layer symbology", "palette": "Project palette", "manual": "Manual",
            "auto": "Automatic (no match)"}.get(s, s)


def group_colors(rows, color_field, group_field, scheme, area_field):
    """Colour for each `group_field` value (e.g. SUBZONA rows on slides/pivots):
    the colour of its dominant `color_field` value by area."""
    if color_field == group_field:
        return {k: v["hex"] for k, v in scheme.items()}
    acc = defaultdict(lambda: defaultdict(float))
    for r in rows:
        g, c = r.get(group_field), r.get(color_field)
        if g not in (None, "") and c not in (None, ""):
            try:
                a = float(r.get(area_field) or 0)
            except (TypeError, ValueError):
                a = 0.0
            acc[str(g).strip()][str(c).strip()] += a or 1e-9
    out = {}
    for g, d in acc.items():
        best = max(d.items(), key=lambda t: t[1])[0]
        if best in scheme:
            out[g] = scheme[best]["hex"]
    return out


def build(rows, fields, field, cfg, lib, layer_scheme=None, palette=None, group_field="SUBZONA",
          area_field="Area_sqm"):
    """Everything the outputs need: dict(field, scheme, groups, info)."""
    if field not in fields:
        field = group_field
    vals = sorted({str(r.get(field)).strip() for r in rows if r.get(field) not in (None, "")})
    info = resolve(vals, field, cfg, lib, layer_scheme, palette)
    return dict(field=field, scheme={k: v["hex"] for k, v in info.items()}, info=info,
                groups=group_colors(rows, field, group_field, info, area_field))
