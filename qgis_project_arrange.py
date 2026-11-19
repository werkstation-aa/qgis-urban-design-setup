"""
QGIS Project Arrange - sort existing layers into the project setup tree
=======================================================================

Run from the QGIS Python Console (Plugins > Python Console > Show Editor >
Open Script > Run). Written for QGIS 3.40 LTR.

Companion to qgis_project_setup.py: that script builds the tree and adds the
datasets; this one tidies a project whose layers are already loaded.

What it does
------------
 1. Reads the name of every layer and group in the project and works out its
    type of information from THEMES below (toll road, substation, coal ...)
 2. Moves each layer to its group. Point layers go to Nodes, at the top of
    the tree, so they draw above everything else.
 3. Builds the setup tree around them: groups, dividers, basemaps last
 4. Sorts each group: points, lines, polygons, rasters, and within each of
    those by type of information, in THEMES order
 5. Applies the setup script's QML styles to the matching layers
 6. Removes groups and old dividers that are left empty

Before arranging, it asks which layer is the project boundary (every run,
Cancel skips this). That layer goes first in Boundaries, is saved as the
"Project Boundary" spatial bookmark, and once everything is in place the map
zooms to it, with the scale rounded up to a golden number so it all fits.

Layers no theme matches stay in the setup group they are in. Elsewhere they go
to "Unsorted" (below the thematic groups); add a keyword to THEMES and re-run.

Nothing is removed from the project, and a layer you could see before stays
visible, a hidden one hidden.

Styles
------
The QML styles are read from the same "styles" folder as the setup script
(STYLES_DIR), so both scripts always use the same ones. A layer is styled once; its previous style is kept
as "Original" in the layer's right-click menu > Styles, next to the setup
styles. Layers named like planned or under-construction features ("Rencana",
"Konstruksi", "Struktur", ...) and LRT/MRT layers keep their own style. Labels
that need a field the layer does not have are switched off.

Set DRY_RUN = True to only print what would happen.
"""

import os
import re
import tempfile
import xml.etree.ElementTree as ET  # parses only the bundled QML style files

from qgis.core import (
    QgsApplication,
    QgsBookmark,
    QgsCoordinateTransform,
    QgsExpression,
    QgsLayerTree,
    QgsProject,
    QgsRasterLayer,
    QgsReferencedRectangle,
    QgsVectorLayer,
    QgsWkbTypes,
)
from qgis.PyQt.QtWidgets import QInputDialog
from qgis.utils import iface


# --------------------------------------------------------------------------
# Configuration (edit here)
# --------------------------------------------------------------------------
DRY_RUN = False             # True: print the plan, change nothing
POINTS_TO_NODES = True      # every point layer goes to Nodes (see NODE_EXEMPT)
SORT_GROUPS = True          # sort each group by geometry, then type of information
APPLY_STYLES = True
RESTYLE_EXISTING = False    # True: style again layers styled by an earlier run

# QML styles shared with the setup script.
STYLES_DIR = "styles"  # folder name (looked up next to this script) or full path

SEPARATOR = "-" * 60

# Same tree as qgis_project_setup.py, top to bottom.
ROOT_GROUPS = [
    "Top Overlays", "Nodes", SEPARATOR,
    "Boundaries", SEPARATOR,
    "Enclaves", "Contour", "Disaster Risk", SEPARATOR,
    "Connectivity", "Phasing", "Design", "Client", SEPARATOR,
    "Infrastructure", "Landcover", "Waterbody", "Topography", SEPARATOR,
    "Orthophoto LIDAR",
]
NODES_GROUP = "Nodes"
NODE_EXEMPT = {"Top Overlays", "Contour", "Topography"}   # their points stay
UNSORTED_GROUP = "Unsorted"
UNSORTED_AFTER = "Topography"       # Unsorted sits below this group
BOUNDARY_NAME = "Project Boundary"  # bookmark name, and the setup script's layer
BOUNDARY_GROUP = "Boundaries"       # the chosen boundary layer goes first here
BOUNDARY_STYLE = ["Base_Project Boundary.qml"]

# Golden scales 1:1000, 1:2000, 1:2500, 1:4000, 1:5000, 1:8000, 1:10000 ...
# Same series as qgis_layout_automation.py and the setup script.
GOLDEN_MANTISSAS = (1.0, 2.0, 2.5, 4.0, 5.0, 8.0)

# Layers the setup script adds and styles itself; left as they are here.
SETUP_LAYER_NAMES = {
    BOUNDARY_NAME, "Train Stations", "Airports", "Seaports",
    "Kabupaten/Kota (Garis)", "Toll Roads", "Main Roads", "Railway",
    "Local Roads (OSM)", "River Ordo 1", "River Ordo 2", "River Ordo 3", "River Ordo 4",
}

# Names with these words keep their own style (planned, under construction,
# spatial-plan structure, light rail).
KEEP_OWN_STYLE = r"\b(rencana|usulan|planned|proposed|konstruksi|construction|struktur|lrt|mrt|brt)\b"

TOLL = ["Base_Connectivity_Toll Road_Black.qml", "Base_Connectivity_Toll Road_Red.qml"]
MAIN_ROAD = ["Base_Connectivity_Main Road_Grey.qml", "Base_Connectivity_Main Road_Orange.qml"]

# Types of information: (name, group, keywords, QML styles).
# A layer gets the first theme whose keywords match its name, so specific
# themes come before general ones. The order is also the order inside each
# group. Names are compared in lower case, with _ - . / ( ) read as spaces and
# CamelCase split ("GarduInduk" -> "gardu induk"). \b marks a whole word.
# The first QML is applied, others are added as alternative styles.
THEMES = [
    ("Labels", "Top Overlays", r"\b(top overlays?|labels?|annotations?|anotasi)\b", []),
    ("Project boundary", "Boundaries",
     r"\b(project boundary|site boundary|batas tapak|delineasi tapak|deliniasi tapak)\b",
     ["Base_Project Boundary.qml"]),
    ("Disaster risk", "Disaster Risk",
     r"\b(risk|risiko|inarisk|hazard|bahaya|rawan|kerawanan|krb)\b", []),
    ("Contours", "Contour", r"\b(contours?|kontur)\b", []),
    ("Phasing", "Phasing", r"\b(phas(e|es|ing)|tahap(an)?)\b", []),
    ("Client", "Client", r"\b(client|klien)\b", []),
    ("Design", "Design",
     r"\b(design|desain|masterplan|master plan|site ?plan|block ?plan|plots?|kavling|massing)\b", []),
    ("Orthophoto", "Orthophoto LIDAR", r"\b(ortho\w*|foto udara|aerial|drone|uav)\b", []),
    # Transport
    ("Seaports", "Connectivity", r"\b(seaports?|ports?|pelabuhan|dermaga|harbou?rs?)\b",
     ["Base_POI_Seaport.qml"]),
    ("Airports", "Connectivity", r"\b(airports?|bandara|bandar udara|airfields?)\b",
     ["Base_POI_Airport.qml"]),
    ("Train stations", "Connectivity", r"\b(stasiun|stations?)\b",
     ["Base_POI_Train Stations.qml"]),
    ("Toll gates", "Connectivity", r"\b(gerbang tol|toll gates?|gt)\b", []),
    ("Terminals", "Connectivity", r"\b(terminals?|halte|bus stops?)\b", []),
    ("Toll roads", "Connectivity",
     r"\b(tol|toll|jalan bebas hambatan|expressways?|freeways?)\b", TOLL),
    ("Main roads", "Connectivity",
     r"\b(jalan nasional|jalan provinsi|arteri|arterial|kolektor|collector"
     r"|main roads?|national roads?|highways?)\b", MAIN_ROAD),
    ("Local roads", "Connectivity",
     r"\b(local roads?|jalan lokal|jalan lingkungan|jalan desa|osm roads?)\b",
     ["Base_Connectivity_Local Road_Purple.qml"]),
    ("Railways", "Connectivity", r"\b(rel|railways?|rail|kereta|ka|lrt|mrt)\b",
     ["Base_Connectivity_Railway.qml"]),
    ("Roads and transport", "Connectivity",
     r"\b(roads?|jalan|jln|streets?|transport\w*|transit|brt|connectivity|konektivitas)\b", []),
    # Settlement, industry, utilities, resources
    ("Urban structure", "Landcover",
     r"\b(perkotaan|urban|pusat kegiatan|pkn|pkw|pkl|permukiman|pemukiman|settlements?)\b", []),
    ("Industry", "Infrastructure",
     r"\b(industri\w*|industrial|kawasan ekonomi|kek|pergudangan|warehous\w*)\b", []),
    ("Energy", "Infrastructure",
     r"\b(energi|energy|gardu|substations?|pembangkit|plt\w*|power|listrik|electric\w*"
     r"|sutet|sutt|transmisi|transmission|pln)\b", []),
    ("Water and sanitation", "Infrastructure",
     r"\b(spam|pdam|air bersih|air minum|water supply|ipal|iplt|tpa|sanitasi|sanitation"
     r"|sewer\w*|drain\w*|limbah|waste)\b", []),
    ("Utilities", "Infrastructure",
     r"\b(infra\w*|utilit\w*|pipa|pipes?|telekomunikasi|telecom\w*|bts)\b", []),
    ("Mining and resources", "Enclaves",
     r"\b(batubara|coal|gas|metana|migas|minyak|oil|tambang|mining|mineral\w*)\b", []),
    ("Permits and concessions", "Enclaves",
     r"\b(wiup\w*|iup\w*|izin|konsesi|concessions?|hgu|permits?)\b", []),
    ("Protected land", "Enclaves",
     r"\b(enclaves?|enklaf|lsd|lahan sawah|lp2b|kawasan lindung|hutan lindung|konservasi"
     r"|conservation|protected|cagar)\b", []),
    ("Spatial plan", "Landcover",
     r"\b(rtrw\w*|rdtr\w*|pola ruang|struktur ruang|spatial plan|zonasi|zoning)\b", []),
    # Physical
    ("Rivers", "Waterbody", r"\b(rivers?|sungai|ordo|streams?|kali)\b", ["Base_Waterbody_LN.qml"]),
    ("Water bodies", "Waterbody",
     r"\b(lakes?|danau|situ|waduk|reservoirs?|waterbod\w*|badan air|coast\w*|pantai|laut|sea)\b",
     []),
    ("Land cover", "Landcover",
     r"\b(land ?cover|land ?use|tutupan|penutupan|penggunaan lahan|pl klhk|vegetasi"
     r"|vegetation|hutan|forest)\b", []),
    ("Topography", "Topography",
     r"\b(topo\w*|dem|dsm|dtm|elevation|elevasi|slope|lereng|kemiringan|hillshade|terrain|lidar)\b",
     []),
    ("Administrative boundaries", "Boundaries",
     r"\b(adm|admin\w*|batas|boundar\w*|provinsi|province|kabupaten|kota|regency|kecamatan"
     r"|district|desa|kelurahan|village|wilayah)\b", ["Base_ADM Boundary_LN.qml"]),
]

# Rasters no theme matches stay loose below the last divider, as basemaps:
# web tiles (XYZ, WMS, WMTS) at the bottom, other rasters above "Orthophoto LIDAR".
BASEMAP = "(basemap)"
OTHER_RASTER = "(raster basemap)"

ORIGINAL_STYLE = "Original"
STYLE_MARK = "project_arrange/style"
BOUNDARY_MARK = "project_arrange/boundary"   # project property: last chosen layer id

boundary_id = None    # id of the layer chosen as project boundary in this run


# --------------------------------------------------------------------------
# Deciding where things go
# --------------------------------------------------------------------------
COMPILED_THEMES = [re.compile(pattern) for _, _, pattern, _ in THEMES]
KEEP_OWN_STYLE_RE = re.compile(KEEP_OWN_STYLE)
SETUP_GROUPS = [name for name in ROOT_GROUPS if name != SEPARATOR]


def normalise(name):
    text = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", name)          # GarduInduk
    text = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1 \2", text)       # KAStation
    return re.sub(r"[_\-./()\[\],:]+", " ", text).lower()


def theme_of(name):
    """Index into THEMES, or None."""
    text = normalise(name)
    for index, pattern in enumerate(COMPILED_THEMES):
        if pattern.search(text):
            return index
    return None


def theme_group(name):
    index = theme_of(name)
    return THEMES[index][1] if index is not None else None


def is_separator(name):
    return re.fullmatch(r"\s*[-–—_=]{3,}\s*", name) is not None


def is_point_layer(layer):
    return isinstance(layer, QgsVectorLayer) and layer.geometryType() == QgsWkbTypes.PointGeometry


def layer_destination(node, current=None):
    """Group name, BASEMAP, OTHER_RASTER or UNSORTED_GROUP.

    current is the setup group the layer is in now, if any.
    """
    if boundary_id and node.layerId() == boundary_id:
        return BOUNDARY_GROUP
    group = theme_group(node.name())
    layer = node.layer()
    if (POINTS_TO_NODES and is_point_layer(layer)
            and group not in NODE_EXEMPT and current not in NODE_EXEMPT):
        return NODES_GROUP
    if group:
        return group
    if isinstance(layer, QgsRasterLayer):
        return BASEMAP if layer.providerType() == "wms" else OTHER_RASTER
    return current or UNSORTED_GROUP


def plan(root):
    """[(node, destination)] for everything that should move, and the groups
    that will be emptied (to remove afterwards, deepest last)."""
    moves, emptied = [], []
    kept = set()

    def already_there(parent, destination):
        if destination in (BASEMAP, OTHER_RASTER):
            return parent is root
        return parent.name() == destination and parent.parent() is root

    def visit(parent, current=None, forced=False):
        """current: the setup group these children belong to, if any.
        forced: parent is a copy of a setup group to merge into current."""
        for child in parent.children():
            if QgsLayerTree.isLayer(child):
                destination = layer_destination(child, current)
                if not already_there(parent, destination):
                    moves.append((child, destination))
                continue
            name = child.name()
            if parent is root and name in SETUP_GROUPS and name not in kept:
                kept.add(name)
                visit(child, current=name)                  # re-check its layers
            elif parent is root and (is_separator(name) or name == UNSORTED_GROUP):
                emptied.append(child)
                visit(child)
            elif parent is root and name in SETUP_GROUPS:
                emptied.append(child)                       # duplicate: merge it
                visit(child, current=name, forced=True)
            elif current and not forced:
                group = theme_group(name)                   # subgroup in a setup group
                if group and group != current:
                    moves.append((child, group))
            elif forced:
                moves.append((child, current))
            elif name in SETUP_GROUPS:
                emptied.append(child)                       # nested copy: merge it
                visit(child, current=name, forced=True)
            else:
                group = theme_group(name)
                if group:
                    moves.append((child, group))            # move the whole group
                else:
                    emptied.append(child)                   # look inside
                    visit(child)

    visit(root)
    return moves, emptied


# --------------------------------------------------------------------------
# Moving tree nodes
# --------------------------------------------------------------------------
def move_node(node, parent, index):
    """Move node to parent at index. Returns the node now in the tree.

    The copy is inserted before the original is removed, so the layer never
    leaves the tree and QGIS never drops it from the project.
    """
    copy = node.clone()
    parent.insertChildNode(index, copy)
    node.parent().removeChildNode(node)
    return copy


def root_group(root, name):
    for child in root.children():
        if QgsLayerTree.isGroup(child) and child.name() == name:
            return child
    group = root.addGroup(name)
    group.setExpanded(False)
    return group


def apply_move(root, node, destination):
    was_visible = node.isVisible()
    if destination in (BASEMAP, OTHER_RASTER):
        target = root
    else:
        target = root_group(root, destination)
    if node.parent() is target:
        return node
    moved = move_node(node, target, len(target.children()))
    # Keep what was visible on the map visible, and what was hidden hidden
    if was_visible and not target.isVisible():
        target.setItemVisibilityChecked(True)
    if not was_visible and moved.isVisible():
        moved.setItemVisibilityChecked(False)
    return moved


def reorder(parent, desired):
    """Put the nodes in desired first in parent, in that order."""
    for index, node in enumerate(desired):
        if parent.children()[index] is not node:
            move_node(node, parent, index)


def sort_key(node, position):
    """(geometry, type of information, original position)"""
    theme = theme_of(node.name())
    theme = len(THEMES) if theme is None else theme
    if QgsLayerTree.isGroup(node):
        return (4, theme, position)
    if node.layerId() == boundary_id or (not boundary_id and node.name() == BOUNDARY_NAME):
        return (-1, 0, position)
    layer = node.layer()
    if isinstance(layer, QgsVectorLayer):
        rank = {QgsWkbTypes.PointGeometry: 0,
                QgsWkbTypes.LineGeometry: 1,
                QgsWkbTypes.PolygonGeometry: 2}.get(layer.geometryType(), 3)
        return (rank, theme, position)
    return (5, theme, position)


def sort_group(group):
    children = group.children()
    desired = sorted(children, key=lambda c: sort_key(c, children.index(c)))
    reorder(group, desired)


def arrange_root(root):
    """Order the root: setup groups with dividers, Unsorted, then basemaps."""
    for child in list(root.children()):
        if QgsLayerTree.isGroup(child) and is_separator(child.name()) and not child.children():
            root.removeChildNode(child)

    loose = [c for c in root.children() if QgsLayerTree.isLayer(c)]
    web = [c for c in loose if layer_destination(c) == BASEMAP]
    other = [c for c in loose if c not in web]
    unsorted = [c for c in root.children()
                if QgsLayerTree.isGroup(c) and c.name() == UNSORTED_GROUP]

    desired = []
    for name in ROOT_GROUPS:
        if name == SEPARATOR:
            divider = root.addGroup(SEPARATOR)
            divider.setItemVisibilityChecked(False)
            desired.append(divider)
            continue
        if name == "Orthophoto LIDAR":
            desired += other
        desired.append(root_group(root, name))
        if name == UNSORTED_AFTER:
            desired += unsorted
            # Anything else (a group no theme matched) goes with Unsorted
            desired += [c for c in root.children()
                        if c not in desired and c not in loose and c not in unsorted
                        and QgsLayerTree.isGroup(c) and c.name() not in ROOT_GROUPS]
    desired += web
    desired += [c for c in root.children() if c not in desired]
    reorder(root, desired)


# --------------------------------------------------------------------------
# Styles
# --------------------------------------------------------------------------
QML_GEOMETRY = {
    "0": QgsWkbTypes.PointGeometry,
    "1": QgsWkbTypes.LineGeometry,
    "2": QgsWkbTypes.PolygonGeometry,
}


def style_library():
    """{qml file name: QML text} from the styles folder, and a problem."""
    candidates = [STYLES_DIR]
    if "__file__" in globals():
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       os.path.basename(STYLES_DIR)))
    path = next((p for p in candidates if p and os.path.isdir(p)), None)
    if path is None:
        return {}, "styles folder not found, set STYLES_DIR: %s" % STYLES_DIR
    library = {}
    for name in os.listdir(path):
        if name.lower().endswith(".qml"):
            with open(os.path.join(path, name), encoding="utf-8") as stream:
                library[name] = stream.read()
    if not library:
        return {}, "no .qml files in %s" % path
    return library, None


def style_label(filename):
    """'Base_Connectivity_Toll Road_Red.qml' -> 'Red'."""
    return os.path.splitext(filename)[0].rsplit("_", 1)[-1]


def qml_geometry(qml):
    try:
        node = ET.fromstring(qml).find("layerGeometryType")
    except ET.ParseError:
        return None
    return QML_GEOMETRY.get((node.text or "").strip()) if node is not None else None


def load_qml(layer, qml):
    """Apply QML text to layer. Returns (message, ok) like loadNamedStyle."""
    handle, path = tempfile.mkstemp(suffix=".qml")
    try:
        with os.fdopen(handle, "w", encoding="utf-8") as stream:
            stream.write(qml)
        return layer.loadNamedStyle(path)
    finally:
        os.remove(path)


def drop_broken_labels(layer):
    """Switch labels off when they use a field the layer does not have.
    Returns the missing field names."""
    if not layer.labelsEnabled() or layer.labeling() is None:
        return set()
    try:
        settings = layer.labeling().settings()
    except (AttributeError, TypeError):
        return set()                      # rule-based labels: leave them
    if settings.isExpression:
        needed = QgsExpression(settings.fieldName).referencedColumns()
    else:
        needed = {settings.fieldName}
    fields = {name.lower() for name in layer.fields().names()}
    missing = {f for f in needed if f and not f.startswith("#!") and f.lower() not in fields}
    if missing:
        layer.setLabelsEnabled(False)
    return missing


def styles_for(layer):
    """(QML files, reason) for a layer. Files is empty when it is not styled."""
    if not isinstance(layer, QgsVectorLayer):
        return [], None
    if layer.name() in SETUP_LAYER_NAMES:
        return [], None
    index = theme_of(layer.name())
    if layer.id() == boundary_id:
        files = BOUNDARY_STYLE
    elif index is None or not THEMES[index][3]:
        return [], None
    else:
        files = THEMES[index][3]
    if layer.id() != boundary_id and KEEP_OWN_STYLE_RE.search(normalise(layer.name())):
        return [], "keeps its own style (planned, structure plan or light rail)"
    if not layer.isValid():
        return [], "layer is not available"
    if layer.customProperty(STYLE_MARK) and not RESTYLE_EXISTING:
        return [], None
    return files, None


def apply_styles(layer, files, library, report):
    styles = []
    for filename in files:
        qml = library.get(filename)
        if qml is None:
            report["notes"].append("%s: %s is not in the styles folder" % (layer.name(), filename))
            continue
        geometry = qml_geometry(qml)
        if geometry is not None and geometry != layer.geometryType():
            report["notes"].append("%s: %s is for another geometry type, skipped"
                                   % (layer.name(), filename))
            continue
        styles.append((style_label(filename), qml))
    if not styles:
        return

    # A QML also carries scale visibility, keep the layer's own limits
    scale = (layer.hasScaleBasedVisibility(), layer.minimumScale(), layer.maximumScale())
    manager = layer.styleManager()
    if ORIGINAL_STYLE not in manager.styles():
        manager.renameStyle(manager.currentStyle(), ORIGINAL_STYLE)
    missing = set()
    for label, qml in styles[1:] + styles[:1]:          # default style last
        if label not in manager.styles():
            manager.addStyleFromLayer(label)
        manager.setCurrentStyle(label)
        message, ok = load_qml(layer, qml)
        if not ok:
            report["failed"].append("%s style %s: %s" % (layer.name(), label, message))
        missing |= drop_broken_labels(layer)
        layer.setScaleBasedVisibility(scale[0])
        layer.setMinimumScale(scale[1])
        layer.setMaximumScale(scale[2])
    layer.setCustomProperty(STYLE_MARK, styles[0][0])
    if missing:
        report["notes"].append("%s: labels off, no field %s"
                               % (layer.name(), ", ".join(sorted(missing))))
    report["styled"].append("%s (%s)" % (layer.name(), " / ".join(l for l, _ in styles)))
    layer.triggerRepaint()
    if iface is not None:
        iface.layerTreeView().refreshLayerSymbology(layer.id())


# --------------------------------------------------------------------------
# Project boundary
# --------------------------------------------------------------------------
def boundary_candidates(project):
    """Vector layers to offer, polygons first, and the index to preselect."""
    vectors = [l for l in project.mapLayers().values()
               if isinstance(l, QgsVectorLayer) and l.isValid()]
    vectors.sort(key=lambda l: (l.geometryType() != QgsWkbTypes.PolygonGeometry,
                                l.name().lower()))
    previous = project.readEntry("ProjectArrange", "boundary", "")[0]
    preselect = 0
    for index, layer in enumerate(vectors):
        if layer.id() == previous:
            return vectors, index
        if layer.name() == BOUNDARY_NAME or (
                theme_of(layer.name()) is not None
                and THEMES[theme_of(layer.name())][0] == "Project boundary"):
            preselect = index
    return vectors, preselect


def ask_boundary_layer(project):
    """Layer picked in a dialog, or None if cancelled or nothing to pick."""
    if iface is None:
        return None
    vectors, preselect = boundary_candidates(project)
    if not vectors:
        return None
    labels = ["%s   (%s)" % (l.name(), QgsWkbTypes.geometryDisplayString(l.geometryType()))
              for l in vectors]
    choice, ok = QInputDialog.getItem(
        iface.mainWindow(), "Project boundary",
        "Which layer is the project boundary?\n"
        "It is saved as the '%s' bookmark and the map zooms to it." % BOUNDARY_NAME,
        labels, preselect, False)
    return vectors[labels.index(choice)] if ok else None


def bookmark_boundary(layer):
    """Save a project bookmark on the layer, replacing any of the same name."""
    manager = QgsProject.instance().bookmarkManager()
    for old in manager.bookmarks():
        if old.name() == BOUNDARY_NAME:
            manager.removeBookmark(old.id())
    bookmark = QgsBookmark()
    bookmark.setName(BOUNDARY_NAME)
    bookmark.setExtent(QgsReferencedRectangle(layer.extent(), layer.crs()))
    manager.addBookmark(bookmark)


def golden_scale(scale):
    """Smallest golden scale denominator >= scale, so the boundary always fits."""
    candidates = sorted(m * 10 ** exp for exp in range(0, 10) for m in GOLDEN_MANTISSAS)
    return next((c for c in candidates if c >= scale * (1 - 1e-9)), candidates[-1])


def zoom_to_boundary(layer, canvas):
    """Fit the map to layer, then round the scale up to a golden number."""
    to_canvas = QgsCoordinateTransform(
        layer.crs(), canvas.mapSettings().destinationCrs(), QgsProject.instance())
    canvas.setExtent(to_canvas.transformBoundingBox(layer.extent()))
    scale = golden_scale(canvas.scale())
    canvas.zoomScale(scale)
    print("Zoomed to %s at 1:%d" % (layer.name(), scale))


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def describe(moves):
    width = max([len(node.name()) for node, _ in moves] + [10])
    for node, destination in moves:
        kind = "group" if QgsLayerTree.isGroup(node) else "layer"
        theme = theme_of(node.name())
        info = THEMES[theme][0] if theme is not None else "-"
        print("  %-*s  %-5s -> %-16s %s" % (width, node.name(), kind, destination, info))


ASK = "ask"   # main(boundary=ASK): pick the boundary in a dialog (console behaviour)


def main(boundary=ASK, report_to="the Python Console"):
    """Arrange the project. The defaults are the console behaviour; the Urban
    Design Setup plugin window passes the layer chosen there (or None) and sets
    DRY_RUN / POINTS_TO_NODES / SORT_GROUPS / APPLY_STYLES / RESTYLE_EXISTING."""
    global boundary_id
    project = QgsProject.instance()
    root = project.layerTreeRoot()
    layer_count = len(project.mapLayers())
    report = {"styled": [], "notes": [], "failed": []}

    if boundary is ASK:
        boundary = ask_boundary_layer(project)
    boundary_id = boundary.id() if boundary is not None else None

    moves, emptied = plan(root)
    style_jobs = []
    if APPLY_STYLES:
        for layer in project.mapLayers().values():
            files, reason = styles_for(layer)
            if files:
                style_jobs.append((layer, files))
            elif reason:
                report["notes"].append("%s: %s" % (layer.name(), reason))

    print("\n=== QGIS project arrange%s ===" % (" (dry run)" if DRY_RUN else ""))
    print("Project boundary: %s" % (boundary.name() if boundary else "none chosen"))
    print("Moves (%d):" % len(moves))
    describe(moves)
    if DRY_RUN:
        print("Styles to apply (%d):" % len(style_jobs))
        for layer, files in style_jobs:
            print("  %s -> %s" % (layer.name(), ", ".join(files)))
        for line in report["notes"]:
            print("  note: " + line)
        return moves

    canvas = iface.mapCanvas() if iface is not None else None
    if canvas is not None:
        canvas.freeze(True)
    try:
        for node, destination in moves:
            apply_move(root, node, destination)
        for group in reversed(emptied):          # deepest first
            if not group.children() and group.parent() is not None:
                group.parent().removeChildNode(group)
        arrange_root(root)
        if SORT_GROUPS:
            for child in root.children():
                if QgsLayerTree.isGroup(child) and not is_separator(child.name()):
                    sort_group(child)
        if style_jobs:
            library, problem = style_library()
            if problem:
                report["failed"].append("Styles: " + problem)
            else:
                for layer, files in style_jobs:
                    apply_styles(layer, files, library, report)
    finally:
        if canvas is not None:
            canvas.freeze(False)

    if boundary is not None:
        bookmark_boundary(boundary)
        project.writeEntry("ProjectArrange", "boundary", boundary.id())
        if canvas is not None:
            QgsApplication.processEvents()     # let QGIS finish its own updates first
            zoom_to_boundary(boundary, canvas)
    if canvas is not None:
        canvas.refresh()

    unsorted =[c.name() for g in root.children()
                if QgsLayerTree.isGroup(g) and g.name() == UNSORTED_GROUP
                for c in g.children()]
    if len(project.mapLayers()) != layer_count:
        report["failed"].append("the project had %d layers and now has %d"
                                % (layer_count, len(project.mapLayers())))
    print("Styled (%d): %s" % (len(report["styled"]), ", ".join(report["styled"]) or "-"))
    for line in report["notes"]:
        print("  note: " + line)
    for line in report["failed"]:
        print("  FAILED: " + line)
    print("Unsorted (%d): %s" % (len(unsorted), ", ".join(unsorted) or "-"))
    if iface is not None:
        text = "Moved %d items, styled %d layers." % (len(moves), len(report["styled"]))
        if unsorted or report["failed"]:
            text += " Some need attention, see %s." % report_to
        iface.messageBar().pushInfo("Project arrange", text)
    return moves


if __name__ != "qgis_project_arrange":
    main()
