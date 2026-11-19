"""
QGIS Project Setup - Layer Tree + Indonesian Data Sources
=========================================================

Run from the QGIS Python Console (Plugins > Python Console > Show Editor >
Open Script > Run). Written for QGIS 3.40 LTR.

What it does
------------
 1. Builds the standard layer tree (groups and separators)
 2. Asks for the project boundary shapefile, adds it to Boundaries and saves
    a spatial bookmark for it (Cancel skips this step)
 3. Adds Esri / Google / OpenStreetMap basemaps
 4. Connects the live datasets into their groups (ArcGIS REST, nothing is
    downloaded), all switched off so opening the project stays fast
 5. Downloads local roads from OpenStreetMap for the site area
 6. Applies the QML styles listed in LAYER_STYLES to the layers it adds
 7. Sets the project CRS to the boundary's CRS and zooms the map to the
    boundary, rounding the scale up to a golden number so it all fits.
    This happens on every run.

Safe to re-run: existing groups and layers are kept, and only what is missing
is added. If a server was down, run the script again to fill the gap. To use
a different boundary, remove the "Project Boundary" layer and re-run.

Local roads need a saved project and a site area, because OSM cannot be
streamed nationally. The project boundary is used as the site area. Without
one, set AOI_BBOX below or zoom the map to the site, then re-run.

Styles
------
The QML styles are plain .qml files in the "styles" folder next to this
script (STYLES_DIR). A layer listed with two QML files gets the first one
applied and the second saved as an alternative, switchable from the layer's
right-click menu > Styles. The layer's own scale limits are kept.

To change a style, re-export the QML from QGIS into the styles folder under
the same file name.

Layers added by an earlier run are not restyled, so manual style edits are
safe. Set RESTYLE_EXISTING = True for one run to style them too.
"""

import os
import tempfile
import xml.etree.ElementTree as ET  # parses only the bundled QML style files
from urllib.parse import quote

from qgis.core import (
    QgsApplication,
    QgsBlockingNetworkRequest,
    QgsBookmark,
    QgsCoordinateReferenceSystem,
    QgsCoordinateTransform,
    QgsDataSourceUri,
    QgsFillSymbol,
    QgsLayerTree,
    QgsProject,
    QgsRasterLayer,
    QgsRectangle,
    QgsReferencedRectangle,
    QgsVectorLayer,
    QgsWkbTypes,
)
from qgis.PyQt.QtCore import QByteArray, QUrl
from qgis.PyQt.QtNetwork import QNetworkRequest
from qgis.PyQt.QtWidgets import QFileDialog
from qgis.utils import iface


# --------------------------------------------------------------------------
# Configuration (edit here)
# --------------------------------------------------------------------------
SEPARATOR = "-" * 60

# Root of the layer tree, top to bottom. Basemaps follow as loose layers.
ROOT_GROUPS = [
    # Point datasets
    "Top Overlays", "Nodes", SEPARATOR,
    # Project boundary
    "Boundaries", SEPARATOR,
    # Overlays
    "Enclaves", "Contour", "Disaster Risk", SEPARATOR,
    # Key planning datasets
    "Connectivity", "Phasing", "Design", "Client", SEPARATOR,
    # Basic physical & infrastructure datasets
    "Infrastructure", "Landcover", "Waterbody", "Topography", SEPARATOR,
    # Basemaps
    "Orthophoto LIDAR",
]
CHECKED_GROUPS = {"Top Overlays"}

BOUNDARY_GROUP = "Boundaries"
BOUNDARY_NAME = "Project Boundary"     # used for the layer and the bookmark
# Fallback when Base_Project Boundary.qml is not found
BOUNDARY_STYLE = {"style": "no", "outline_color": "black", "outline_width": "0.8"}

# Golden scales 1:1000, 1:2000, 1:2500, 1:4000, 1:5000, 1:8000, 1:10000 ...
# Same series as qgis_layout_automation.py.
GOLDEN_MANTISSAS = (1.0, 2.0, 2.5, 4.0, 5.0, 8.0)

# Site area as (west, south, east, north) in degrees, for example
# (106.70, -6.35, 106.95, -6.10). None uses the project boundary, or the
# current map view if there is no boundary.
AOI_BBOX = None
OSM_MAX_AREA_DEG2 = 1.0       # refuse larger areas, Overpass would time out
OSM_FOLDER = "data"           # relative to the saved project file
OSM_HIGHWAY_TYPES = ("trunk|primary|secondary|tertiary|unclassified|"
                     "residential|living_street|service|track")
OVERPASS_URLS = [             # tried in order, public servers are often busy
    "https://overpass-api.de/api/interpreter",
    "https://overpass.private.coffee/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://maps.mail.ru/osm/tools/overpass/api/interpreter",
]
LOCAL_ROADS_NAME = "Local Roads (OSM)"

DEFAULT_BASEMAP = "Google Satellite"   # the only layer switched on
RISK_OPACITY = 0.7
BLANK_PROJECT_CRS = "EPSG:3857"        # only used when there is no boundary
INDONESIA_BBOX = (95.0, -11.0, 141.0, 6.0)

BIG = "https://kspservices.big.go.id/satupeta/rest/services/PUBLIK"
BNPB = "https://gis.bnpb.go.id/server/rest/services"
SIGI = "https://sigi.pu.go.id/serverpu/rest/services/sigi_postgis"

# (group, layer name, url, hide when zoomed out past 1:N, or None)
# Limits are on the layers too heavy to draw nationally.
VECTOR_LAYERS = [
    ("Nodes", "Train Stations", BIG + "/SARANA_PRASARANA/MapServer/7", None),
    ("Nodes", "Airports", BIG + "/SARANA_PRASARANA/MapServer/4", None),
    ("Nodes", "Seaports", BIG + "/SARANA_PRASARANA/MapServer/1", None),

    ("Boundaries", "Provinsi", BIG + "/BATAS_WILAYAH/MapServer/0", None),
    ("Boundaries", "Kabupaten/Kota (Area)", BIG + "/BATAS_WILAYAH/MapServer/2", None),
    ("Boundaries", "Kabupaten/Kota (Garis)", BIG + "/BATAS_WILAYAH/MapServer/1", None),
    ("Boundaries", "Kecamatan", BIG + "/BATAS_WILAYAH/MapServer/3", 2000000),
    ("Boundaries", "Desa", BIG + "/BATAS_WILAYAH/MapServer/4", 500000),

    # BIG toll layer, because the Binamarga service (gisportal.binamarga.pu.go.id
    # Tol/Jalan_Tol_Operasi) fails every query and draws blank maps.
    ("Connectivity", "Toll Roads", BIG + "/SARANA_PRASARANA/MapServer/14", None),
    ("Connectivity", "Main Roads", BIG + "/SARANA_PRASARANA/MapServer/13", None),
    ("Connectivity", "Railway", BIG + "/SARANA_PRASARANA/MapServer/6", None),

    ("Landcover", "Landcover 2024", BNPB + "/thematic/PL_KLHK_2024/MapServer/0", 250000),
    ("Landcover", "Landcover 2020", BNPB + "/thematic/PL_KLHK_2020/MapServer/0", 250000),
    ("Landcover", "Landcover 2017", BNPB + "/thematic/PL_KLHK_2017/MapServer/0", 250000),

    ("Waterbody", "River Ordo 1", SIGI + "/ast_sungai_ordo1/MapServer/0", None),
    ("Waterbody", "River Ordo 2", SIGI + "/ast_sungai_ordo2/MapServer/0", 1000000),
    ("Waterbody", "River Ordo 3", SIGI + "/ast_sungai_ordo3/MapServer/0", 250000),
    ("Waterbody", "River Ordo 4", SIGI + "/ast_sungai_ordo4/MapServer/0", 250000),
]

# InaRISK rasters, drawn by the server: (layer name, service slug)
RISK_GROUP = "Disaster Risk"
RISK_LAYERS = [
    ("Risk - Flood", "banjir"),
    ("Risk - Flash Flood", "banjir_bandang"),
    ("Risk - Earthquake", "gempabumi"),
    ("Risk - Tsunami", "tsunami"),
    ("Risk - Landslide", "tanah_longsor"),
    ("Risk - Liquefaction", "likuefaksi"),
    ("Risk - Extreme Waves and Abrasion", "gelombang_ekstrim_dan_abrasi"),
    ("Risk - Extreme Weather", "cuaca_ekstrim"),
    ("Risk - Forest Fire", "kebakaran_hutan_dan_lahan"),
    ("Risk - Drought", "kekeringan"),
    ("Risk - Covid19", "covid19"),
]

# XYZ tiles: (name, tile url, max zoom)
BASEMAPS = [
    ("Google Satellite", "https://www.google.cn/maps/vt?lyrs=s@189&gl=cn&x={x}&y={y}&z={z}", 20),
    ("Esri Imagery", "https://server.arcgisonline.com/arcgis/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}", 19),
    ("OpenStreetMap", "https://tile.openstreetmap.org/{z}/{x}/{y}.png", 19),
]

RESTYLE_EXISTING = False   # True also styles layers kept from an earlier run
STYLES_DIR = "styles"      # QML folder: name (looked up next to this script) or full path

# Layer name: QML files (embedded at the bottom of this script). The first is
# applied, any others are saved as alternative styles named after the last
# word of the file name ("Red"). A style made for another geometry type
# (points, lines, polygons) is skipped. Add a layer here and put its QML in
# the styles folder.
# Not used yet: Base_Waterbody_AR.qml, Base_Connectivity_Secondary Road_Orange.qml
LAYER_STYLES = {
    BOUNDARY_NAME: ["Base_Project Boundary.qml"],
    "Train Stations": ["Base_POI_Train Stations.qml"],
    "Airports": ["Base_POI_Airport.qml"],
    "Seaports": ["Base_POI_Seaport.qml"],
    "Kabupaten/Kota (Garis)": ["Base_ADM Boundary_LN.qml"],
    "Toll Roads": ["Base_Connectivity_Toll Road_Black.qml",
                   "Base_Connectivity_Toll Road_Red.qml"],
    "Main Roads": ["Base_Connectivity_Main Road_Grey.qml",
                   "Base_Connectivity_Main Road_Orange.qml"],
    "Railway": ["Base_Connectivity_Railway.qml"],
    LOCAL_ROADS_NAME: ["Base_Connectivity_Local Road_Purple.qml"],
    "River Ordo 1": ["Base_Waterbody_LN.qml"],
    "River Ordo 2": ["Base_Waterbody_LN.qml"],
    "River Ordo 3": ["Base_Waterbody_LN.qml"],
    "River Ordo 4": ["Base_Waterbody_LN.qml"],
}


# --------------------------------------------------------------------------
# Layer tree helpers
# --------------------------------------------------------------------------
def find_groups(root, name):
    return [c for c in root.children() if QgsLayerTree.isGroup(c) and c.name() == name]


def find_layer_index(parent, name):
    """Index of the direct child layer called name, or None."""
    for index, child in enumerate(parent.children()):
        if QgsLayerTree.isLayer(child) and child.name() == name:
            return index
    return None


def build_groups(root):
    """Create missing root groups in order. Returns the index after the last."""
    cursor = 0
    seen = {}
    for name in ROOT_GROUPS:
        nth = seen.get(name, 0)
        seen[name] = nth + 1
        matches = find_groups(root, name)
        if nth < len(matches):
            cursor = root.children().index(matches[nth]) + 1
            continue
        group = root.insertGroup(cursor, name)
        group.setExpanded(False)
        group.setItemVisibilityChecked(name in CHECKED_GROUPS)
        cursor += 1
    return cursor


def place_layer(parent, layer, position, checked=False):
    """Add layer to the project under parent, without drawing it."""
    QgsProject.instance().addMapLayer(layer, False)
    node = parent.insertLayer(min(position, len(parent.children())), layer)
    node.setItemVisibilityChecked(checked)
    node.setExpanded(False)


# --------------------------------------------------------------------------
# QML styles
# --------------------------------------------------------------------------
QML_GEOMETRY = {
    "0": QgsWkbTypes.PointGeometry,
    "1": QgsWkbTypes.LineGeometry,
    "2": QgsWkbTypes.PolygonGeometry,
}


def styles_dir():
    """STYLES_DIR as given, or the folder of that name next to this script."""
    if os.path.isdir(STYLES_DIR):
        return STYLES_DIR
    if "__file__" in globals():
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), os.path.basename(STYLES_DIR))
    return STYLES_DIR


def style_qml(filename):
    """QML text of a file in the styles folder, or None."""
    path = os.path.join(styles_dir(), filename)
    if not os.path.isfile(path):
        return None
    with open(path, encoding="utf-8") as stream:
        return stream.read()


def style_label(filename):
    """'Base_Connectivity_Toll Road_Red.qml' -> 'Red'."""
    return os.path.splitext(filename)[0].rsplit("_", 1)[-1]


def style_fits(layer, qml):
    """False when the QML was saved from a different geometry type."""
    if not isinstance(layer, QgsVectorLayer):
        return True
    try:
        node = ET.fromstring(qml).find("layerGeometryType")
    except ET.ParseError:
        return True                   # loadNamedStyle reports the problem
    expected = QML_GEOMETRY.get((node.text or "").strip()) if node is not None else None
    return expected is None or expected == layer.geometryType()


def load_qml(layer, qml):
    """Apply QML text to layer. Returns (message, ok) like loadNamedStyle."""
    handle, path = tempfile.mkstemp(suffix=".qml")
    try:
        with os.fdopen(handle, "w", encoding="utf-8") as stream:
            stream.write(qml)
        return layer.loadNamedStyle(path)
    finally:
        os.remove(path)


def style_layer(layer, report):
    """Apply the QML styles listed for the layer. Returns True if styled."""
    styles = []
    for filename in LAYER_STYLES.get(layer.name(), []):
        qml = style_qml(filename)
        if qml is None:
            problem = "%s is not in %s" % (filename, styles_dir())
        elif not style_fits(layer, qml):
            problem = "%s was saved from a different geometry type" % filename
        else:
            styles.append((style_label(filename), qml))
            continue
        report["skipped"].append("%s style: %s" % (layer.name(), problem))
    if not styles:
        return False

    # A QML also carries scale visibility, keep the layer's own limits
    scale = (layer.hasScaleBasedVisibility(), layer.minimumScale(), layer.maximumScale())
    manager = layer.styleManager()
    alternatives = styles[1:]
    if alternatives:
        # Park the current style so styles from an earlier run can be replaced
        manager.renameStyle(manager.currentStyle(), "_setup")
        for label, _ in styles:
            if label in manager.styles():
                manager.removeStyle(label)
        for label, qml in alternatives:
            message, ok = load_qml(layer, qml)
            if ok:
                manager.addStyleFromLayer(label)
            else:
                report["failed"].append("%s style %s: %s" % (layer.name(), label, message))

    default_label, default_qml = styles[0]
    message, ok = load_qml(layer, default_qml)
    if alternatives:
        manager.renameStyle(manager.currentStyle(), default_label)
    layer.setScaleBasedVisibility(scale[0])
    layer.setMinimumScale(scale[1])
    layer.setMaximumScale(scale[2])
    if not ok:
        report["failed"].append("%s style: %s" % (layer.name(), message))
        return False
    report["styled"].append(layer.name())
    return True


def restyle_kept(node, report):
    """Style a layer from an earlier run, only when RESTYLE_EXISTING is on."""
    layer = node.layer()
    if not RESTYLE_EXISTING or layer is None or not style_layer(layer, report):
        return
    layer.triggerRepaint()
    if iface is not None:
        iface.layerTreeView().refreshLayerSymbology(layer.id())


# --------------------------------------------------------------------------
# Project boundary
# --------------------------------------------------------------------------
def golden_scale(scale):
    """Smallest golden scale denominator >= scale, so the boundary always fits."""
    candidates = sorted(m * 10 ** exp for exp in range(0, 10) for m in GOLDEN_MANTISSAS)
    return next((c for c in candidates if c >= scale * (1 - 1e-9)), candidates[-1])


def ask_boundary_path():
    """Path chosen in the file dialog, or "" if cancelled."""
    if iface is None:
        return ""
    path, _ = QFileDialog.getOpenFileName(
        iface.mainWindow(), "Select the project boundary shapefile",
        QgsProject.instance().homePath(), "Shapefiles (*.shp)")
    return path


def add_boundary(root, report, boundary_path=None):
    """Add the chosen shapefile to its group. Returns the new layer, or None.

    boundary_path: None asks in a file dialog, "" skips, else the shapefile to use.
    """
    group = find_groups(root, BOUNDARY_GROUP)[0]
    existing = find_layer_index(group, BOUNDARY_NAME)
    if existing is not None:
        report["kept"].append(BOUNDARY_NAME)
        restyle_kept(group.children()[existing], report)
        return None

    path = ask_boundary_path() if boundary_path is None else boundary_path
    if not path:
        report["skipped"].append(BOUNDARY_NAME + ": no shapefile selected")
        return None

    layer = QgsVectorLayer(path, BOUNDARY_NAME, "ogr")
    if not layer.isValid():
        problem = "could not open " + path
    elif not layer.crs().isValid():
        problem = "the shapefile has no CRS (missing .prj file)"
    elif layer.extent().isEmpty():
        problem = "the shapefile has no usable extent"
    else:
        problem = None
    if problem:
        report["failed"].append("%s: %s" % (BOUNDARY_NAME, problem))
        return None

    if (not style_layer(layer, report)
            and layer.geometryType() == QgsWkbTypes.PolygonGeometry):
        layer.renderer().setSymbol(QgsFillSymbol.createSimple(BOUNDARY_STYLE))
    place_layer(group, layer, 0, checked=True)
    group.setItemVisibilityChecked(True)
    report["added"].append(BOUNDARY_NAME)
    return layer


def add_boundary_bookmark(layer):
    """Save a project bookmark on the layer, replacing any of the same name."""
    manager = QgsProject.instance().bookmarkManager()
    for old in manager.bookmarks():
        if old.name() == BOUNDARY_NAME:
            manager.removeBookmark(old.id())
    bookmark = QgsBookmark()
    bookmark.setName(BOUNDARY_NAME)
    bookmark.setExtent(QgsReferencedRectangle(layer.extent(), layer.crs()))
    manager.addBookmark(bookmark)


def zoom_to_boundary(layer, canvas):
    """Fit the map to layer, then round the scale up to a golden number."""
    to_canvas = QgsCoordinateTransform(
        layer.crs(), canvas.mapSettings().destinationCrs(), QgsProject.instance())
    canvas.setExtent(to_canvas.transformBoundingBox(layer.extent()))
    scale = golden_scale(canvas.scale())
    canvas.zoomScale(scale)
    print("Map scale set to 1:%d" % scale)


# --------------------------------------------------------------------------
# Layer builders
# --------------------------------------------------------------------------
def make_vector(name, url, scale_limit=None):
    uri = QgsDataSourceUri()
    uri.setParam("url", url)
    layer = QgsVectorLayer(uri.uri(False), name, "arcgisfeatureserver")
    if layer.isValid() and scale_limit:
        layer.setScaleBasedVisibility(True)
        layer.setMinimumScale(scale_limit)
    return layer


def make_risk(name, slug):
    uri = QgsDataSourceUri()
    uri.setParam("url", "%s/inarisk/layer_risiko_%s/ImageServer" % (BNPB, slug))
    uri.setParam("layer", "")
    uri.setParam("format", "png32")
    layer = QgsRasterLayer(uri.uri(False), name, "arcgismapserver")
    if layer.isValid():
        layer.setOpacity(RISK_OPACITY)
    return layer


def make_basemap(name, url, zmax):
    uri = QgsDataSourceUri()
    uri.setParam("type", "xyz")
    uri.setParam("url", url)
    uri.setParam("zmin", "0")
    uri.setParam("zmax", str(zmax))
    return QgsRasterLayer(bytes(uri.encodedUri()).decode("utf-8"), name, "wms")


def add_datasets(root, report):
    """Add VECTOR_LAYERS and RISK_LAYERS to their groups."""
    position = {}
    boundary_index = find_layer_index(find_groups(root, BOUNDARY_GROUP)[0], BOUNDARY_NAME)
    if boundary_index is not None:
        position[BOUNDARY_GROUP] = boundary_index + 1   # keep it on top

    def add(group_name, name, build):
        group = find_groups(root, group_name)[0]
        slot = position.get(group_name, 0)
        existing = find_layer_index(group, name)
        if existing is not None:
            position[group_name] = existing + 1
            report["kept"].append(name)
            restyle_kept(group.children()[existing], report)
            return
        print("Connecting %s ..." % name)
        layer = build()
        if not layer.isValid():
            report["failed"].append("%s: %s" % (name, layer.error().summary() or "no response"))
            return
        style_layer(layer, report)
        place_layer(group, layer, slot)
        position[group_name] = slot + 1
        report["added"].append(name)

    for group, name, url, limit in VECTOR_LAYERS:
        add(group, name, lambda n=name, u=url, m=limit: make_vector(n, u, m))

    for name, slug in RISK_LAYERS:
        add(RISK_GROUP, name, lambda n=name, s=slug: make_risk(n, s))


def add_basemaps(root, cursor, report):
    for name, url, zmax in BASEMAPS:
        existing = find_layer_index(root, name)
        if existing is not None:
            cursor = existing + 1
            report["kept"].append(name)
            continue
        layer = make_basemap(name, url, zmax)
        if not layer.isValid():
            report["failed"].append("%s: %s" % (name, layer.error().summary()))
            continue
        place_layer(root, layer, cursor, checked=(name == DEFAULT_BASEMAP))
        cursor += 1
        report["added"].append(name)


# --------------------------------------------------------------------------
# Local roads (OpenStreetMap)
# --------------------------------------------------------------------------
def site_bbox():
    """(west, south, east, north) in degrees: AOI_BBOX, boundary, or map view."""
    if AOI_BBOX:
        return AOI_BBOX
    project = QgsProject.instance()
    boundaries = project.mapLayersByName(BOUNDARY_NAME)
    if boundaries:
        source_crs, extent = boundaries[0].crs(), boundaries[0].extent()
    elif iface is not None:
        canvas = iface.mapCanvas()
        source_crs, extent = canvas.mapSettings().destinationCrs(), canvas.extent()
    else:
        return None
    to_wgs84 = QgsCoordinateTransform(
        source_crs, QgsCoordinateReferenceSystem("EPSG:4326"), project)
    box = to_wgs84.transformBoundingBox(extent)
    return box.xMinimum(), box.yMinimum(), box.xMaximum(), box.yMaximum()


def download_osm_roads(bbox, path):
    """Save OSM roads inside bbox to path. Returns an error message or None."""
    west, south, east, north = bbox
    query = ('[out:xml][timeout:180];way["highway"~"^(%s)$"](%f,%f,%f,%f);'
             "(._;>;);out body;" % (OSM_HIGHWAY_TYPES, south, west, north, east))
    body = QByteArray(("data=" + quote(query, safe="")).encode("utf-8"))
    error = None
    for url in OVERPASS_URLS:
        request = QNetworkRequest(QUrl(url))
        request.setHeader(QNetworkRequest.ContentTypeHeader,
                          "application/x-www-form-urlencoded")
        blocking = QgsBlockingNetworkRequest()
        if blocking.post(request, body) == QgsBlockingNetworkRequest.NoError:
            content = bytes(blocking.reply().content())
            # Busy servers can answer "OK" with an error page or an empty result
            if b"<way " in content:
                with open(path, "wb") as handle:
                    handle.write(content)
                return None
            error = "%s returned no roads" % url
        else:
            error = blocking.errorMessage()
        print("Overpass server failed, trying the next: %s" % error)
    return error


def add_local_roads(root, report):
    group = find_groups(root, "Connectivity")[0]
    existing = find_layer_index(group, LOCAL_ROADS_NAME)
    if existing is not None:
        report["kept"].append(LOCAL_ROADS_NAME)
        restyle_kept(group.children()[existing], report)
        return

    project = QgsProject.instance()
    if not project.fileName():
        report["skipped"].append(LOCAL_ROADS_NAME + ": save the project first, then re-run")
        return
    bbox = site_bbox()
    if bbox is None or (bbox[2] - bbox[0]) * (bbox[3] - bbox[1]) > OSM_MAX_AREA_DEG2:
        report["skipped"].append(
            LOCAL_ROADS_NAME + ": zoom to the site or set AOI_BBOX, then re-run")
        return

    folder = os.path.join(project.homePath(), OSM_FOLDER)
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, "local_roads.osm")
    print("Downloading %s ..." % LOCAL_ROADS_NAME)
    error = download_osm_roads(bbox, path)
    if error:
        report["failed"].append("%s: %s" % (LOCAL_ROADS_NAME, error))
        return

    layer = QgsVectorLayer(path + "|layername=lines", LOCAL_ROADS_NAME, "ogr")
    if not layer.isValid() or not layer.setSubsetString('"highway" IS NOT NULL'):
        report["failed"].append(LOCAL_ROADS_NAME + ": no roads in the downloaded file")
        return
    style_layer(layer, report)
    place_layer(group, layer, len(group.children()))
    report["added"].append(LOCAL_ROADS_NAME)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def print_report(report, report_to="the Python Console"):
    print("\n=== QGIS project setup ===")
    print("Added %d, kept %d existing, styled %d" % (
        len(report["added"]), len(report["kept"]), len(report["styled"])))
    for title, key in (("Skipped", "skipped"), ("FAILED", "failed")):
        if report[key]:
            print("%s:" % title)
            for line in report[key]:
                print("  - " + line)
    if iface is not None:
        problems = len(report["failed"]) + len(report["skipped"])
        text = "Added %d layers, styled %d." % (len(report["added"]), len(report["styled"]))
        if problems:
            text += " %d need attention, see %s." % (problems, report_to)
        iface.messageBar().pushInfo("Project setup", text)


def main(boundary_path=None, local_roads=True, report_to="the Python Console"):
    """Run the setup. The defaults are the console behaviour; the Urban Design
    Setup plugin window passes its own choices (and trims VECTOR_LAYERS,
    RISK_LAYERS and BASEMAPS to the datasets ticked there).

    boundary_path: None asks in a file dialog, "" skips, else the shapefile.
    """
    project = QgsProject.instance()
    root = project.layerTreeRoot()
    blank_project = not project.mapLayers()
    report = {"added": [], "kept": [], "styled": [], "skipped": [], "failed": []}

    if blank_project:
        project.setCrs(QgsCoordinateReferenceSystem(BLANK_PROJECT_CRS))

    canvas = iface.mapCanvas() if iface is not None else None
    cursor = build_groups(root)
    boundary = add_boundary(root, report, boundary_path)
    if boundary is not None:
        add_boundary_bookmark(boundary)

    if canvas is not None:
        canvas.freeze(True)
    try:
        add_basemaps(root, cursor, report)
        add_datasets(root, report)
        if local_roads:
            add_local_roads(root, report)
    finally:
        if canvas is not None:
            canvas.freeze(False)

    if canvas is not None:
        # QGIS sets its own CRS and zoom after the first layers of an empty
        # project. Let that finish now, otherwise it overrides what follows.
        QgsApplication.processEvents()

    boundaries = project.mapLayersByName(BOUNDARY_NAME)
    if boundaries:
        project.setCrs(boundaries[0].crs())
        print("Project CRS set to %s" % project.crs().authid())
    elif blank_project:
        project.setCrs(QgsCoordinateReferenceSystem(BLANK_PROJECT_CRS))

    if canvas is not None:
        if boundaries:
            zoom_to_boundary(boundaries[0], canvas)
        elif blank_project:
            to_project = QgsCoordinateTransform(
                QgsCoordinateReferenceSystem("EPSG:4326"), project.crs(), project)
            canvas.setExtent(to_project.transformBoundingBox(QgsRectangle(*INDONESIA_BBOX)))
        canvas.refresh()
    print_report(report, report_to)
    return report


if __name__ != "qgis_project_setup":
    main()
