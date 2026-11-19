"""QGIS side: read parcels from a land-use layer, write the population memory layer."""

from qgis.PyQt.QtCore import QVariant
from qgis.core import (
    QgsDistanceArea,
    QgsFeature,
    QgsField,
    QgsFields,
    QgsProject,
    QgsUnitTypes,
    QgsVectorLayer,
    QgsWkbTypes,
)

from . import pop_core as core

GEOM_AREA = "__geometry__"
POP_FIELDS = [  # name, type, length, precision
    ("Pop_Resi", QVariant.Int, 10, 0),
    ("Pop_Visit", QVariant.Int, 10, 0),
    ("Pop_Total", QVariant.Int, 10, 0),
    ("Pop_Type", QVariant.String, 20, 0),
    ("Pop_Units", QVariant.Double, 12, 1),
    ("Pop_GFA", QVariant.Double, 14, 2),
    ("Pop_GFASrc", QVariant.String, 12, 0),
    ("Pop_Asmp", QVariant.String, 80, 0),
    ("Pop_Method", QVariant.String, 120, 0),
    ("Pop_Status", QVariant.String, 12, 0),
    ("Pop_Source", QVariant.String, 254, 0),
    ("Pop_QA", QVariant.String, 120, 0),
]


def _sqm(da, geom):
    a = da.measureArea(geom)
    try:
        from qgis.core import Qgis
        return da.convertAreaMeasurement(a, Qgis.AreaUnit.SquareMeters)  # 3.30+
    except (ImportError, AttributeError):
        return da.convertAreaMeasurement(a, QgsUnitTypes.AreaSquareMeters)


def collect(layer, class_field, area_field=GEOM_AREA, gfa_field=None, far_field=None, selected=False):
    """-> list of records {fid, cls, area_m2, gfa, gfa_src, qa}."""
    da = QgsDistanceArea()
    da.setSourceCrs(layer.crs(), QgsProject.instance().transformContext())
    da.setEllipsoid(QgsProject.instance().ellipsoid() or "EPSG:7030")
    feats = layer.getSelectedFeatures() if selected and layer.selectedFeatureCount() else layer.getFeatures()
    out = []
    for f in feats:
        if area_field and area_field != GEOM_AREA:
            area = core._num(f[area_field]) or 0.0
        else:
            area = _sqm(da, f.geometry()) if f.hasGeometry() else 0.0
        gfa, src, qa = core.resolve_gfa(f[gfa_field] if gfa_field else None,
                                        f[far_field] if far_field else None, area)
        v = f[class_field]
        cls = "(empty)" if v is None or str(v).strip() in ("", "NULL") else str(v)
        out.append(dict(fid=f.id(), cls=cls, area_m2=area, gfa=gfa, gfa_src=src, qa=qa))
    return out


LANDUSE_COLORS = {  # fallback when the layer is not categorised on the class field (Export Land Use palette)
    "industrial": "ADDDF7", "hotel": "CD0000", "commercial": "FF6D6D", "mice": "833232",
    "residential": "FFEF82", "dormitory": "FFDF00", "office": "BA3B5D", "port": "544CC3",
    "public amenity": "9963CC", "golf": "346E2B", "green": "58BA48", "road": "FFFFFF", "utility": "999999",
}


def class_colors(layer, class_field):
    """{class value: 'RRGGBB'} from the layer symbology when it is categorised on the class field;
    other classes fall back to the land-use palette in color_for()."""
    out = {}
    try:
        r = layer.renderer()
        if r is not None and r.type() == "categorizedSymbol" and r.classAttribute() == class_field:
            for cat in r.categories():
                if cat.symbol() is not None:
                    out[str(cat.value())] = cat.symbol().color().name()[1:].upper()
        # single symbol / graduated / rule-based: no per-class colour -> palette in color_for()
    except Exception:
        pass
    return out


def color_for(cls, colors):
    if cls in colors:
        return colors[cls]
    low = str(cls).lower()
    for k, v in LANDUSE_COLORS.items():
        if k in low:
            return v
    return None


def population_values(rec, assignment, lib):
    prof = lib["by_id"].get(assignment["profile"], lib["by_id"]["NONE"])
    res = core.compute(prof["method"], assignment["params"], prof["pop_type"], rec["area_m2"], rec["gfa"])
    resi, visit = int(round(res["resi"])), int(round(res["visit"]))
    qa = "; ".join(x for x in (rec["qa"], res["msg"]) if x)
    return {
        "Pop_Resi": resi, "Pop_Visit": visit, "Pop_Total": resi + visit,
        "Pop_Type": prof["pop_type"] if (resi or visit) else "-",
        "Pop_Units": round(res["units"], 1) if res["units"] is not None else None,
        "Pop_GFA": round(rec["gfa"], 2) if rec["gfa"] else None, "Pop_GFASrc": rec["gfa_src"] or None,
        "Pop_Asmp": prof["name"][:80], "Pop_Method": core.formula_text(prof["method"], assignment["params"])[:120],
        "Pop_Status": prof["status"], "Pop_Source": core.source_text(lib, prof)[:254], "Pop_QA": qa[:120] or None,
    }


def build_memory_layer(layer, records, assignments, lib, name=None):
    """Copy of the parcels (all, or the ones in records) with the population fields. Existing
    Pop_* fields of the source are overwritten, not duplicated."""
    pop_names = [p[0] for p in POP_FIELDS]
    fields = QgsFields()
    keep = []
    for fld in layer.fields():
        if fld.name() in pop_names:
            continue
        fields.append(QgsField(fld))
        keep.append(fld.name())
    for n, t, ln, pr in POP_FIELDS:
        fields.append(QgsField(n, t, len=ln, prec=pr))

    geom = QgsWkbTypes.displayString(layer.wkbType())
    mem = QgsVectorLayer(f"{geom}?crs={layer.crs().authid() or 'EPSG:4326'}",
                         name or f"{layer.name()} - population", "memory")
    if not layer.crs().authid():
        mem.setCrs(layer.crs())
    prov = mem.dataProvider()
    prov.addAttributes(fields.toList())
    mem.updateFields()

    by_fid = {r["fid"]: r for r in records}
    feats, table = [], []
    for f in layer.getFeatures(list(by_fid)):
        rec = by_fid.get(f.id())
        if rec is None:
            continue
        vals = population_values(rec, assignments.get(rec["cls"]) or core.default_assignment(rec["cls"], lib), lib)
        nf = QgsFeature(mem.fields())
        nf.setGeometry(f.geometry())
        attrs = [f[n] for n in keep] + [vals[n] for n in pop_names]
        nf.setAttributes(attrs)
        feats.append(nf)
        row = {n: _py(f[n]) for n in keep}
        row.update(vals)
        table.append(row)
    prov.addFeatures(feats)
    mem.updateExtents()
    return mem, keep + pop_names, table


def _py(v):
    """QVariant / Qt types -> plain Python for openpyxl."""
    if v is None:
        return None
    try:
        from qgis.PyQt.QtCore import QDate, QDateTime
        if isinstance(v, QDateTime):
            return v.toPyDateTime()
        if isinstance(v, QDate):
            return v.toPyDate()
    except ImportError:
        pass
    if isinstance(v, QVariant):
        return None if v.isNull() else v.value()
    if isinstance(v, (int, float, str, bool)):
        return v
    return str(v)
