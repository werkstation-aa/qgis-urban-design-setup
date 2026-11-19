"""Read a QGIS vector layer + its symbology into the plain data vpe_core writes.
Shared by the Processing algorithm and the preview window, so both give the same
workbook."""

from collections import Counter, defaultdict

from qgis.PyQt.QtCore import QDate, QDateTime, QTime, QVariant
from qgis.core import (QgsDistanceArea, QgsExpressionContextUtils, QgsFeatureRequest, QgsRenderContext,
                       QgsUnitTypes, QgsWkbTypes)

from . import vpe_core as core

AGGS = ["sum", "mean", "min", "max", "count"]
AGG_OPTIONS = ["Sum", "Average", "Min", "Max", "Count"]
# (label, kind, divisor from m² / m)
MEASURES = [("Auto (polygons: area in ha, lines: length in m, points: none)", None, None),
            ("None", None, None),
            ("Area (m²)", "area", 1.0), ("Area (ha)", "area", 1e4), ("Area (km²)", "area", 1e6),
            ("Length (m)", "length", 1.0), ("Length (km)", "length", 1e3)]
MEASURE_COL = {("area", 1.0): "Area_m2", ("area", 1e4): "Area_ha", ("area", 1e6): "Area_km2",
               ("length", 1.0): "Length_m", ("length", 1e3): "Length_km"}
PIVOT_OPTIONS = ["Excel pivot table if Microsoft Excel is installed, else static table",
                 "Static table (values, no Excel needed)",
                 "Excel pivot table (fail if Excel is not available)"]
PIVOT_KEYS = ["auto", "static", "excel"]
SYMBOLOGY = "<symbology>"  # row / column token for the legend-class column


class CollectError(Exception):
    pass


def py(v):
    """QGIS attribute value -> plain Python value for Excel."""
    if v is None or (isinstance(v, QVariant) and v.isNull()):
        return None
    if isinstance(v, QDateTime):
        return v.toString("yyyy-MM-dd HH:mm:ss")
    if isinstance(v, QDate):
        return v.toString("yyyy-MM-dd")
    if isinstance(v, QTime):
        return v.toString("HH:mm:ss")
    if isinstance(v, (bool, int, float, str)):
        return v
    return str(v)


def unique_name(name, taken):
    """Excel table headers are case-insensitive: 'Area_ha' clashes with 'Area_Ha'."""
    low = {str(t).lower() for t in taken if t}
    out, i = name, 2
    while out.lower() in low:
        out, i = f"{name}_{i}", i + 1
    return out


def symbol_hex(sym):
    """Fill colour of a symbol; the outline colour when the fill is transparent."""
    if sym is None:
        return None
    c = sym.color()
    if c.alpha() == 0 and sym.symbolLayerCount():
        try:
            c = sym.symbolLayer(0).strokeColor()
        except Exception:
            pass
    return c.name()[1:].upper() if c.isValid() and c.alpha() > 0 else None


def symbology_field(layer):
    """The plain attribute a categorised / graduated renderer classifies on, else None."""
    r = layer.renderer()
    if r is not None and r.type() in ("categorizedSymbol", "graduatedSymbol"):
        f = r.classAttribute()
        if layer.fields().indexOf(f) >= 0:
            return f
    return None


def legend_order(layer, field):
    r = layer.renderer()
    if r is not None and r.type() == "categorizedSymbol" and r.classAttribute() == field:
        return [py(c.value()) for c in r.categories() if py(c.value()) not in (None, "")]
    return []


def resolve_measure(layer, measure_index):
    """(kind, divisor, warning) for MEASURES[measure_index] on this layer."""
    _, kind, div = MEASURES[measure_index]
    gtype = QgsWkbTypes.geometryType(layer.wkbType())
    if measure_index == 0:  # auto
        kind, div = {QgsWkbTypes.PolygonGeometry: ("area", 1e4),
                     QgsWkbTypes.LineGeometry: ("length", 1.0)}.get(gtype, (None, None))
    if (kind == "area" and gtype != QgsWkbTypes.PolygonGeometry) or \
            (kind == "length" and gtype != QgsWkbTypes.LineGeometry):
        return None, None, "The geometry measure does not fit this layer's geometry; skipped."
    return kind, div, None


def collect(layer, selected=False, row_fields=(), col_field=None, value_fields=(), agg="sum",
            measure_index=0, add_count=True, percent=True, ellipsoid="", transform_context=None,
            feedback=None):
    """-> dict(fields, rows, spec, colors, orders, title, note, warnings, sym_col).

    row_fields / col_field may hold SYMBOLOGY for the legend-class column. Empty
    row_fields = the field the layer is styled on, else the symbology class."""
    warnings = []
    kind, div, warn = resolve_measure(layer, measure_index)
    if warn:
        warnings.append(warn)

    attr = [f.name() for f in layer.fields()]
    taken = set(attr)
    sym_col = unique_name("Symbology", taken)
    taken.add(sym_col)
    count_col = unique_name("Features", taken) if add_count else None
    taken.add(count_col)
    measure_col = unique_name(MEASURE_COL[(kind, div)], taken) if kind else None
    fields = attr + [sym_col] + [c for c in (measure_col, count_col) if c]

    def real(f):
        return sym_col if f == SYMBOLOGY else f

    row_fields = [real(f) for f in row_fields if f]
    col_field = real(col_field) if col_field else None
    if not row_fields:
        row_fields = [symbology_field(layer) or sym_col]
    measures = [(f, agg) for f in value_fields]
    if measure_col:
        measures.append((measure_col, "sum"))
    if count_col:
        measures.append((count_col, "sum"))
    if not measures:
        raise CollectError("Nothing to summarise: pick value fields, a geometry measure or the feature count.")

    da = QgsDistanceArea()
    if transform_context is not None:
        da.setSourceCrs(layer.crs(), transform_context)
    da.setEllipsoid(ellipsoid or layer.crs().ellipsoidAcronym() or "NONE")
    renderer = layer.renderer().clone() if layer.renderer() is not None else None
    rctx = QgsRenderContext()
    rctx.expressionContext().appendScopes(QgsExpressionContextUtils.globalProjectLayerScopes(layer))
    labels = {}
    if renderer is not None:
        renderer.startRender(rctx, layer.fields())
        labels = {it.ruleKey(): it.label() for it in renderer.legendSymbolItems()}
    legend_keys = list(labels)

    request = QgsFeatureRequest()
    if selected:
        request.setFilterFids(layer.selectedFeatureIds())
    total = (layer.selectedFeatureCount() if selected else layer.featureCount()) or 1
    rows, feat_hex = [], []
    try:
        for i, f in enumerate(layer.getFeatures(request)):
            if feedback is not None and feedback.isCanceled():
                break
            r = {n: py(v) for n, v in zip(attr, f.attributes())}
            hx, lab = None, None
            if renderer is not None:
                rctx.expressionContext().setFeature(f)
                # plural call: rule-based renderers return None from originalSymbolForFeature
                syms = renderer.originalSymbolsForFeature(f, rctx)
                hx = next((h for h in map(symbol_hex, syms) if h), None)
                keys = renderer.legendKeysForFeature(f, rctx)
                lab = next((labels[k] for k in legend_keys if k in keys and labels[k]), None)
                if lab is None and renderer.type() == "singleSymbol":
                    lab = layer.name()  # single symbol has no legend label
            r[sym_col] = lab
            if measure_col:
                g = f.geometry()
                m = None
                if g is not None and not g.isEmpty():
                    if kind == "area":
                        m = _to_m2(da, da.measureArea(g)) / div
                    else:
                        m = _to_m(da, da.measureLength(g)) / div
                r[measure_col] = m
            if count_col:
                r[count_col] = 1
            rows.append(r)
            feat_hex.append(hx)
            if feedback is not None:
                feedback.setProgress(80 * (i + 1) / total)
    finally:
        if renderer is not None:
            renderer.stopRender(rctx)

    # colour per value: the most common feature colour
    colored = list(dict.fromkeys(row_fields + ([col_field] if col_field else []) + [sym_col]))
    colors, orders = {}, {}
    for fld in colored:
        votes = defaultdict(Counter)
        for r, hx in zip(rows, feat_hex):
            if hx:
                votes[core.key(r.get(fld))][hx] += 1
        colors[fld] = {v: c.most_common(1)[0][0] for v, c in votes.items()}
        orders[fld] = legend_order(layer, fld)
    orders[sym_col] = [labels[k] for k in labels if labels[k]] if renderer is not None else []

    spec = dict(row_fields=row_fields, col_field=col_field, measures=measures, percent=percent)
    note = (f"Layer '{layer.name()}', renderer '{renderer.type() if renderer else 'none'}'. "
            "Each value takes the colour of the features carrying it (the most common one).")
    return dict(fields=fields, rows=rows, spec=spec, colors=colors, orders=orders,
                title=f"{layer.name().upper()} - PIVOT", note=note, warnings=warnings, sym_col=sym_col)


def _unit(name_new, name_old):
    """Qgis.AreaUnit / DistanceUnit (3.30+) or QgsUnitTypes (older)."""
    try:
        from qgis.core import Qgis
        enum, member = name_new.split(".")
        return getattr(getattr(Qgis, enum), member)
    except AttributeError:
        return getattr(QgsUnitTypes, name_old)


def _to_m2(da, v):
    return da.convertAreaMeasurement(v, _unit("AreaUnit.SquareMeters", "AreaSquareMeters"))


def _to_m(da, v):
    return da.convertLengthMeasurement(v, _unit("DistanceUnit.Meters", "DistanceMeters"))
