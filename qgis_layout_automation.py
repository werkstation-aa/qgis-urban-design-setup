"""
QGIS Layout Automation - Spatial Bookmark x Map Themes
=======================================================

Run from the QGIS Python Console (Plugins > Python Console > Show Editor >
Open Script > Run). Tested against the QGIS 3.x PyQGIS API (3.16+).

Workflow
--------
 1. Setup Spatial Bookmark      -> pick a bookmark (project or user bookmarks)
 2. Define Themes to export     -> tick the map themes to export
 3. Create Layout, image size   -> choose an output pixel size (page = px @ 300 dpi),
                                   or set MANUAL_PAGE_SIZE below in mm / px
 4. Add Map, full-bleed         -> map item covers the whole page, no frame
 5. Add Scalebar                -> position (corner), style, segments, units, segment size;
                                   drawn on the page preview exactly as exported
 6. Add North arrow             -> left of the scalebar (default) or its own corner
 7. Create pages                -> one page per selected theme
 8. Round map scale             -> golden numbers 1:1000, 1:2000, 1:2500, 1:4000,
                                   1:5000, 1:8000, 1:10000, ... or a manual scale;
                                   a live preview shows the bookmark boundary on the page
 9. Export as image             -> 300 dpi PNG, one file per theme
10. Select export folder        -> folder picker in the dialog

The layout is kept in the project's Layout Manager so it can be fine-tuned
and re-exported manually.
"""

import functools
import math
import os
import re

from qgis.core import (
    Qgis,
    QgsApplication,
    QgsCoordinateTransform,
    QgsLayoutExporter,
    QgsLayoutItem,
    QgsLayoutItemMap,
    QgsLayoutItemPage,
    QgsLayoutItemPicture,
    QgsLayoutItemScaleBar,
    QgsLayoutPoint,
    QgsLayoutSize,
    QgsPrintLayout,
    QgsProject,
    QgsRectangle,
    QgsScaleBarSettings,
    QgsSymbolLayerUtils,
    QgsPathResolver,
    QgsTextBufferSettings,
    QgsTextFormat,
    QgsUnitTypes,
)
from qgis.PyQt.QtCore import Qt, QByteArray, QRectF, QSize, QTimer
from qgis.PyQt.QtGui import QColor, QFont, QGuiApplication, QImage, QPainter, QPen, QPixmap
from qgis.PyQt.QtSvg import QSvgRenderer
from qgis.PyQt.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)
from qgis.utils import iface


# --------------------------------------------------------------------------
# Defaults / style configuration (edit here)
# --------------------------------------------------------------------------
DPI = 300
LAYOUT_NAME = "Auto_Bookmark_Themes"

# Exported PNG name (without .png). Fields: {theme} {bookmark} {layout} {scale} {page}
#   "{theme}"                  -> Landuse.png
#   "{bookmark}_{theme}"       -> Site_Landuse.png
#   "{theme}_1-{scale}"        -> Landuse_1-8000.png
# Characters not allowed in Windows file names are replaced with "_".
FILENAME_TEMPLATE = "{theme}"

# Page size presets (width, height, unit). Paper sizes are exact millimetres,
# screen sizes are exact pixels; the other unit is derived at DPI.
IMAGE_PRESETS = {
    "A4 Landscape (297 x 210 mm)": (297, 210, "mm"),
    "A4 Portrait (210 x 297 mm)": (210, 297, "mm"),
    "A3 Landscape (420 x 297 mm)": (420, 297, "mm"),
    "A3 Portrait (297 x 420 mm)": (297, 420, "mm"),
    "Full HD (1920 x 1080 px)": (1920, 1080, "px"),
    "4K UHD (3840 x 2160 px)": (3840, 2160, "px"),
    "Square (3000 x 3000 px)": (3000, 3000, "px"),
    "Custom": None,
}

# Manual page size, typed here in the script, in millimetres or pixels:
#   MANUAL_PAGE_SIZE = (420, 297, "mm")     # A3 landscape in mm
#   MANUAL_PAGE_SIZE = (500, 300, "mm")     # any custom sheet
#   MANUAL_PAGE_SIZE = (5000, 3000, "px")   # exact output pixels
# When set, it appears first in the preset list ("Manual (script)") and is
# selected by default. Set to None to start on the first preset above.
MANUAL_PAGE_SIZE = None


def manual_page_size():
    """MANUAL_PAGE_SIZE as (w, h, "mm" | "px"), kept exact, or None."""
    if not MANUAL_PAGE_SIZE:
        return None
    w, h, unit = MANUAL_PAGE_SIZE
    unit = unit.lower()
    if unit not in ("mm", "px"):
        raise ValueError('MANUAL_PAGE_SIZE unit must be "mm" or "px", got %r' % unit)
    return float(w), float(h), unit

# Golden scale mantissas per decade -> 1:1000, 1:2000, 1:2500, 1:4000, 1:5000, 1:8000, 1:10000 ...
GOLDEN_MANTISSAS = (1.0, 2.0, 2.5, 4.0, 5.0, 8.0)

CORNERS = ["Bottom left", "Bottom right", "Top left", "Top right"]

ITEM_MARGIN_MM = 6.0          # distance of scalebar / north arrow from page edge

SCALEBAR_STYLES = ["Single Box", "Double Box", "Line Ticks Middle", "Line Ticks Down",
                   "Line Ticks Up", "Stepped Line", "Hollow", "Numeric"]
SCALEBAR_FONT = "Arial"
SCALEBAR_FONT_SIZE_PT = 8
SCALEBAR_HEIGHT_MM = 2.0
SCALEBAR_MIN_WIDTH_MM = 25.0
SCALEBAR_MAX_WIDTH_MM = 45.0
SCALEBAR_SEGMENTS = 4          # segments right of 0
SCALEBAR_SEGMENTS_LEFT = 0     # subdivisions of the extra segment left of 0 (0 = none)
SCALEBAR_UNITS = "auto"        # "auto" (km from SCALEBAR_KM_FROM_SCALE), "m" or "km"
SCALEBAR_SEGMENT_SIZE = None   # None = auto, bar fits MIN/MAX width; or fixed size per
                               # segment in scalebar units, e.g. 100 (m) or 0.5 (km)
SCALEBAR_KM_FROM_SCALE = 25000  # use km labels at/above this scale denominator

SCALEBAR_TEXT_HALO = True      # white buffer around labels (no background box)

# QGIS bundled SVG library name, or a full path to any .svg file
NORTH_ARROW_SVG = "arrows/NorthArrow_01.svg"
NORTH_ARROW_HEIGHT_RATIO = 1.0      # drawn arrow height / scalebar height (1.0 = flush)
NORTH_ARROW_BESIDE_SCALEBAR = True  # north arrow sits left of the scalebar, same corner
NORTH_ARROW_GAP_MM = 3.0            # space between north arrow and scalebar
NORTH_ARROW_FILL = QColor(0, 0, 0)          # black fill
NORTH_ARROW_STROKE = QColor(255, 255, 255)  # white outline, keeps it visible on dark maps
NORTH_ARROW_STROKE_WIDTH_MM = 0.12  # thin edge; thicker eats into the black ring


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def px_to_mm(px, dpi=DPI):
    return px / dpi * 25.4


def mm_to_px(mm, dpi=DPI):
    return mm / 25.4 * dpi


def golden_series():
    """All golden scale denominators, 1 ... 8e9, ascending."""
    return sorted(m * 10 ** exp for exp in range(0, 10) for m in GOLDEN_MANTISSAS)


def golden_scale(scale, mode="up"):
    """Snap a scale denominator to the golden series.

    mode='up'      -> smallest golden scale >= scale (bookmark extent always fits)
    mode='nearest' -> closest golden scale on a log basis
    """
    if scale <= 0:
        return scale
    candidates = golden_series()
    if mode == "up":
        for c in candidates:
            if c >= scale * (1 - 1e-9):
                return c
        return candidates[-1]
    return min(candidates, key=lambda c: abs(math.log(c / scale)))


def distance_unit(name):
    """Distance units across QGIS versions (Qgis.DistanceUnit since 3.30)."""
    try:
        return getattr(Qgis.DistanceUnit, name)
    except AttributeError:
        return getattr(QgsUnitTypes, "Distance" + name)


def scalebar_unit(scale, units="auto"):
    """'m' or 'km' for the scalebar labels."""
    if units in ("m", "km"):
        return units
    return "km" if scale >= SCALEBAR_KM_FROM_SCALE else "m"


def default_scalebar_options():
    return {"segments": SCALEBAR_SEGMENTS, "segments_left": SCALEBAR_SEGMENTS_LEFT,
            "units": SCALEBAR_UNITS, "segment_size": SCALEBAR_SEGMENT_SIZE}


def safe_name(text):
    return re.sub(r'[\\/:*?"<>|]+', "_", text).strip() or "untitled"


def all_bookmarks():
    """Project bookmarks first, then user (global) bookmarks."""
    out = []
    for bm in QgsProject.instance().bookmarkManager().bookmarks():
        out.append(("[Project] " + bm.name(), bm))
    for bm in QgsApplication.bookmarkManager().bookmarks():
        out.append(("[User] " + bm.name(), bm))
    return out


def corner_anchor(corner, page_w, page_h, margin):
    """Return (reference point, x, y) in page-local mm for a corner placement."""
    if corner == "Bottom left":
        return QgsLayoutItem.LowerLeft, margin, page_h - margin
    if corner == "Bottom right":
        return QgsLayoutItem.LowerRight, page_w - margin, page_h - margin
    if corner == "Top left":
        return QgsLayoutItem.UpperLeft, margin, margin
    return QgsLayoutItem.UpperRight, page_w - margin, margin


def move_item(item, ref, x, y, page_index):
    item.setReferencePoint(ref)
    # attemptMove(point, useReferencePoint, includesFrame, page)
    item.attemptMove(QgsLayoutPoint(x, y, QgsUnitTypes.LayoutMillimeters), True, False, page_index)


def place_item(item, corner, page_index, page_w, page_h, margin=ITEM_MARGIN_MM):
    ref, x, y = corner_anchor(corner, page_w, page_h, margin)
    move_item(item, ref, x, y, page_index)


@functools.lru_cache(maxsize=None)
def svg_visible_bounds(svg_path, px=240):
    """Drawn content of an SVG as (left, top, right, bottom) fractions of its box.

    Bundled SVGs pad the artwork inside the viewBox (NorthArrow_01's circle
    fills ~88%), so this lets the visible arrow, not its box, line up with
    the scalebar. QGIS param(...) placeholders are resolved to opaque black.
    """
    with open(svg_path, encoding="utf-8") as f:
        text = f.read()
    text = re.sub(r"param\((fill|outline)\)", "#000000", text)
    text = re.sub(r"param\([a-z-]*opacity\)", "1", text)
    text = re.sub(r"param\(outline-width\)", "0", text)
    renderer = QSvgRenderer(QByteArray(text.encode("utf-8")))
    vb = renderer.viewBoxF()
    w, h = px, max(1, int(round(px * vb.height() / vb.width())))
    img = QImage(w, h, QImage.Format_ARGB32)
    img.fill(Qt.transparent)
    painter = QPainter(img)
    renderer.render(painter)
    painter.end()
    xs, ys = [], []
    for y in range(h):
        for x in range(w):
            if img.pixelColor(x, y).alpha() > 10:
                xs.append(x)
                ys.append(y)
    if not xs:
        return 0.0, 0.0, 1.0, 1.0
    return min(xs) / w, min(ys) / h, (max(xs) + 1) / w, (max(ys) + 1) / h


def drawn_insets(item, dpi=200):
    """Empty space (left, top, right, bottom) in mm between an item's box and its ink.

    The scalebar box reserves room below the bar, so aligning to the box leaves
    the arrow hanging low. This renders the item alone and finds its dark pixels
    (labels + bar outline; the white label halo is ignored).
    """
    layout = item.layout()
    hidden = []
    for other in layout.items():
        if (isinstance(other, QgsLayoutItem) and not isinstance(other, QgsLayoutItemPage)
                and other is not item and other.isVisible()):
            other.setVisibility(False)
            hidden.append(other)
    try:
        box = item.sceneBoundingRect()
        img = QgsLayoutExporter(layout).renderRegionToImage(box, QSize(), dpi)
    finally:
        for other in hidden:
            other.setVisibility(True)
    w, h = img.width(), img.height()
    xs, ys = [], []
    for y in range(h):
        for x in range(w):
            c = img.pixelColor(x, y)
            if c.alpha() > 10 and c.lightness() < 160:
                xs.append(x)
                ys.append(y)
    if not xs:
        return 0.0, 0.0, 0.0, 0.0
    sx, sy = box.width() / w, box.height() / h
    return (min(xs) * sx, min(ys) * sy, (w - max(xs) - 1) * sx, (h - max(ys) - 1) * sy)


def size_arrow_to_scalebar(arrow, scalebar):
    """Resize the north arrow so its drawn artwork is as tall as the scalebar's ink."""
    _, t, _, b = svg_visible_bounds(arrow.picturePath())
    _, sb_top, _, sb_bottom = drawn_insets(scalebar)
    drawn_sb_h = scalebar.sizeWithUnits().height() - sb_top - sb_bottom
    target_h = drawn_sb_h * NORTH_ARROW_HEIGHT_RATIO
    renderer = QSvgRenderer(arrow.picturePath())
    vb = renderer.viewBoxF()
    box_h = target_h / (b - t)
    box_w = box_h * vb.width() / vb.height()
    arrow.attemptResize(QgsLayoutSize(box_w, box_h, QgsUnitTypes.LayoutMillimeters))


def place_arrow_beside_scalebar(scalebar, arrow, corner, page_index, page_w, page_h,
                                margin=ITEM_MARGIN_MM, gap=NORTH_ARROW_GAP_MM):
    """North arrow in the scalebar's corner, scalebar shifted right to make room.

    Left corners:  [margin][arrow][gap][scalebar]
    Right corners:          [arrow][gap][scalebar][margin]
    Positions use what is actually drawn (the arrow's artwork, the scalebar's
    labels and bar), not the padded item boxes, so the margin and gap are
    exact and the arrow's top and bottom line up with the scalebar's.
    """
    l, t, r, b = svg_visible_bounds(arrow.picturePath())
    box_w, box_h = arrow.sizeWithUnits().width(), arrow.sizeWithUnits().height()
    sb_l, sb_t, sb_r, sb_b = drawn_insets(scalebar)
    sb_drawn_w = scalebar.sizeWithUnits().width() - sb_l - sb_r

    if corner.startswith("Bottom"):
        ref_l, ref_r = QgsLayoutItem.LowerLeft, QgsLayoutItem.LowerRight
        ink_y = page_h - margin                 # bottom of the bar
        sb_y, arrow_y = ink_y + sb_b, ink_y + (1 - b) * box_h
    else:
        ref_l, ref_r = QgsLayoutItem.UpperLeft, QgsLayoutItem.UpperRight
        ink_y = margin                          # top of the labels
        sb_y, arrow_y = ink_y - sb_t, ink_y - t * box_h

    if corner.endswith("left"):
        move_item(arrow, ref_l, margin - l * box_w, arrow_y, page_index)
        move_item(scalebar, ref_l, margin + (r - l) * box_w + gap - sb_l, sb_y, page_index)
    else:
        move_item(scalebar, ref_r, page_w - margin + sb_r, sb_y, page_index)
        arrow_drawn_right = page_w - margin - sb_drawn_w - gap
        move_item(arrow, ref_r, arrow_drawn_right + (1 - r) * box_w, arrow_y, page_index)


# --------------------------------------------------------------------------
# Dialog widgets
# --------------------------------------------------------------------------
class GoldenScaleSpinBox(QSpinBox):
    """Scale denominator box: arrows / wheel step through the golden series,
    while any value can still be typed."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setRange(1, 100000000)
        self.setGroupSeparatorShown(True)
        self.setValue(1000)

    def stepBy(self, steps):
        v = self.value()
        if steps > 0:
            bigger = [c for c in golden_series() if c > v]
            if bigger:
                self.setValue(int(bigger[min(steps, len(bigger)) - 1]))
        elif steps < 0:
            smaller = [c for c in golden_series() if c < v]
            if smaller:
                self.setValue(int(smaller[-min(-steps, len(smaller))]))

    def stepEnabled(self):
        return QSpinBox.StepUpEnabled | QSpinBox.StepDownEnabled


class PagePreview(QWidget):
    """Page outline with the bookmark boundary drawn to scale on it.

    The view zooms out when the boundary is larger than the page, so the
    cropped part is visible outside the sheet.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(280, 220)
        self.page_mm = None
        self.boundary_mm = None
        self.decorations = None  # [(QImage, QRectF in page mm)] scalebar + north arrow

    def set_state(self, page_mm, boundary_mm):
        self.page_mm, self.boundary_mm = page_mm, boundary_mm
        self.update()

    def set_decorations(self, items):
        self.decorations = items
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.fillRect(self.rect(), QColor(70, 70, 70))
        if not self.page_mm:
            return
        pw, ph = self.page_mm
        bw, bh = self.boundary_mm or (0, 0)
        pad = 14
        k = min((self.width() - 2 * pad) / max(pw, bw), (self.height() - 2 * pad) / max(ph, bh))
        cx, cy = self.width() / 2, self.height() / 2

        page = QRectF(cx - pw * k / 2, cy - ph * k / 2, pw * k, ph * k)
        p.setPen(QPen(QColor(160, 160, 160), 1))
        p.setBrush(QColor(255, 255, 255))
        p.drawRect(page)

        if self.boundary_mm:
            bound = QRectF(cx - bw * k / 2, cy - bh * k / 2, bw * k, bh * k)
            pen = QPen(QColor(220, 40, 40), 1.6, Qt.DashLine)
            p.setPen(pen)
            p.setBrush(QColor(220, 40, 40, 45))
            p.drawRect(bound)

        # scalebar + north arrow, true size and position on the page
        p.setRenderHint(QPainter.SmoothPixmapTransform)
        for img, r in self.decorations or []:
            p.drawImage(QRectF(page.x() + r.x() * k, page.y() + r.y() * k,
                               r.width() * k, r.height() * k), img)

        p.setPen(QColor(90, 90, 90))
        f = p.font()
        f.setPointSizeF(8)
        p.setFont(f)
        p.drawText(page.adjusted(5, 3, -5, -3), Qt.AlignLeft | Qt.AlignTop,
                   "Page %.0f x %.0f mm" % (pw, ph))
        p.setPen(QColor(200, 40, 40))
        p.drawText(page.adjusted(5, 3, -5, -3), Qt.AlignRight | Qt.AlignTop, "- - Bookmark boundary")
        p.end()


# --------------------------------------------------------------------------
# Dialog (steps 1, 2, 3, 5, 6, 8, 10)
# --------------------------------------------------------------------------
class LayoutAutomationDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Layout Automation - Bookmark x Themes")
        self.setMinimumSize(420, 320)       # small enough for any screen; content scrolls
        self.setSizeGripEnabled(True)
        outer = QVBoxLayout(self)
        # everything except the OK / Cancel buttons scrolls, so the dialog can be
        # made smaller than its contents without anything going off-screen
        self._content = QWidget()
        cols = QHBoxLayout(self._content)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)
        scroll.setWidget(self._content)
        outer.addWidget(scroll, 1)
        root = QVBoxLayout()            # left column: inputs
        side = QVBoxLayout()            # right column: scale + page preview
        cols.addLayout(root, 1)
        cols.addLayout(side, 1)

        # 1. Spatial bookmark
        self.bookmarks = all_bookmarks()
        self.cmb_bookmark = QComboBox()
        for label, _ in self.bookmarks:
            self.cmb_bookmark.addItem(label)
        grp_bm = QGroupBox("1. Spatial bookmark")
        lay_bm = QVBoxLayout(grp_bm)
        lay_bm.addWidget(self.cmb_bookmark)
        root.addWidget(grp_bm)

        # 2. Themes
        self.lst_themes = QListWidget()
        for theme in QgsProject.instance().mapThemeCollection().mapThemes():
            it = QListWidgetItem(theme)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked)
            self.lst_themes.addItem(it)
        btn_all = QPushButton("Select all")
        btn_none = QPushButton("Select none")
        btn_all.clicked.connect(lambda: self._set_all_themes(Qt.Checked))
        btn_none.clicked.connect(lambda: self._set_all_themes(Qt.Unchecked))
        grp_th = QGroupBox("2. Map themes to export (one page each)")
        lay_th = QVBoxLayout(grp_th)
        lay_th.addWidget(self.lst_themes)
        row = QHBoxLayout()
        row.addWidget(btn_all)
        row.addWidget(btn_none)
        lay_th.addLayout(row)
        root.addWidget(grp_th)

        # 3. Layout + image size
        grp_ly = QGroupBox("3. Layout and image size (@ %d dpi)" % DPI)
        form = QFormLayout(grp_ly)
        self.txt_layout = QLineEdit(LAYOUT_NAME)
        self.presets = {}
        manual = manual_page_size()
        if manual:
            self.presets["Manual (script) %g x %g %s" % manual] = manual
        self.presets.update(IMAGE_PRESETS)
        self.cmb_preset = QComboBox()
        self.cmb_preset.addItems(list(self.presets.keys()))
        self.cmb_unit = QComboBox()
        self.cmb_unit.addItems(["mm", "px"])
        self._unit = "mm"
        self._exact = next(v for v in self.presets.values() if v)
        self.spn_w = QDoubleSpinBox()
        self.spn_h = QDoubleSpinBox()
        self._configure_size_boxes()
        self.lbl_size = QLineEdit()
        self.lbl_size.setReadOnly(True)
        self.cmb_preset.currentTextChanged.connect(self._apply_preset)
        self.cmb_unit.currentTextChanged.connect(self._change_unit)
        self.spn_w.valueChanged.connect(self._size_edited)
        self.spn_h.valueChanged.connect(self._size_edited)
        form.addRow("Layout name", self.txt_layout)
        form.addRow("Preset", self.cmb_preset)
        form.addRow("Unit", self.cmb_unit)
        form.addRow("Width", self.spn_w)
        form.addRow("Height", self.spn_h)
        form.addRow("Equals", self.lbl_size)
        root.addWidget(grp_ly)

        # 5./6. Decorations
        grp_dec = QGroupBox("5-6. Scalebar and north arrow")
        form_dec = QFormLayout(grp_dec)
        self.cmb_sb_corner = QComboBox()
        self.cmb_sb_corner.addItems(CORNERS)
        self.cmb_sb_corner.setCurrentText("Bottom left")
        self.cmb_sb_style = QComboBox()
        self.cmb_sb_style.addItems(SCALEBAR_STYLES)
        self.cmb_na_corner = QComboBox()
        self.cmb_na_corner.addItems(CORNERS)
        self.cmb_na_corner.setCurrentText("Top right")
        self.chk_beside = QCheckBox("North arrow left of scalebar (same corner)")
        self.chk_beside.toggled.connect(lambda on: self.cmb_na_corner.setEnabled(not on))
        self.chk_beside.setChecked(NORTH_ARROW_BESIDE_SCALEBAR)
        self.cmb_na_corner.setEnabled(not NORTH_ARROW_BESIDE_SCALEBAR)
        self.spn_seg = QSpinBox()
        self.spn_seg.setRange(1, 20)
        self.spn_seg.setValue(SCALEBAR_SEGMENTS)
        self.spn_seg_left = QSpinBox()
        self.spn_seg_left.setRange(0, 10)
        self.spn_seg_left.setValue(SCALEBAR_SEGMENTS_LEFT)
        self.spn_seg_left.setSpecialValueText("None")
        self.spn_seg_left.setToolTip("Adds one segment left of 0, split into this many parts")
        self.cmb_sb_units = QComboBox()
        self.cmb_sb_units.addItem("Auto (km from 1:%s)" % f"{SCALEBAR_KM_FROM_SCALE:,}", "auto")
        self.cmb_sb_units.addItem("Metres", "m")
        self.cmb_sb_units.addItem("Kilometres", "km")
        self.cmb_sb_units.setCurrentIndex(max(0, self.cmb_sb_units.findData(SCALEBAR_UNITS)))
        self.cmb_seg_mode = QComboBox()
        self.cmb_seg_mode.addItem("Auto (bar %g-%g mm)" % (SCALEBAR_MIN_WIDTH_MM, SCALEBAR_MAX_WIDTH_MM), "fit")
        self.cmb_seg_mode.addItem("Fixed", "fixed")
        self.spn_seg_size = QDoubleSpinBox()
        self.spn_seg_size.setDecimals(3)
        self.spn_seg_size.setRange(0.001, 1000000)
        self.spn_seg_size.setValue(SCALEBAR_SEGMENT_SIZE or 100)
        self.cmb_seg_mode.setCurrentIndex(1 if SCALEBAR_SEGMENT_SIZE else 0)
        self.spn_seg_size.setEnabled(bool(SCALEBAR_SEGMENT_SIZE))
        self.cmb_seg_mode.currentIndexChanged.connect(self._segment_mode_changed)
        row_seg_size = QHBoxLayout()
        row_seg_size.addWidget(self.cmb_seg_mode)
        row_seg_size.addWidget(self.spn_seg_size)
        form_dec.addRow("Scalebar position", self.cmb_sb_corner)
        form_dec.addRow("Scalebar style", self.cmb_sb_style)
        form_dec.addRow("Segments", self.spn_seg)
        form_dec.addRow("Left segment", self.spn_seg_left)
        form_dec.addRow("Units", self.cmb_sb_units)
        form_dec.addRow("Segment size", row_seg_size)
        form_dec.addRow(self.chk_beside)
        form_dec.addRow("North arrow position", self.cmb_na_corner)
        root.addWidget(grp_dec)

        # the decoration preview is rebuilt shortly after the last change, not per keystroke
        self._deco_timer = QTimer(self)
        self._deco_timer.setSingleShot(True)
        self._deco_timer.setInterval(150)
        self._deco_timer.timeout.connect(self._render_decorations)
        self._last_deco_info = None
        for sig in (self.cmb_sb_corner.currentIndexChanged, self.cmb_sb_style.currentIndexChanged,
                    self.cmb_na_corner.currentIndexChanged, self.chk_beside.toggled,
                    self.spn_seg.valueChanged, self.spn_seg_left.valueChanged,
                    self.cmb_sb_units.currentIndexChanged, self.spn_seg_size.valueChanged):
            sig.connect(self._schedule_decorations)

        # 8. Map scale + preview of the bookmark boundary on the page
        grp_sc = QGroupBox("8. Map scale and page preview")
        lay_sc = QVBoxLayout(grp_sc)
        form_sc = QFormLayout()
        self.cmb_round = QComboBox()
        self.cmb_round.addItem("Golden, round up - whole bookmark visible", "up")
        self.cmb_round.addItem("Golden, round to nearest", "nearest")
        self.cmb_round.addItem("Manual scale", "manual")
        self.spn_scale = GoldenScaleSpinBox()
        self.spn_scale.setToolTip("Arrows / mouse wheel step through 1:1000, 1:2000, 1:2500, "
                                  "1:4000, 1:5000, 1:8000 ...; any value can be typed.")
        form_sc.addRow("Scale mode", self.cmb_round)
        form_sc.addRow("Scale  1 :", self.spn_scale)
        lay_sc.addLayout(form_sc)
        self.preview = PagePreview()
        lay_sc.addWidget(self.preview, 1)
        self.lbl_preview = QLabel()
        self.lbl_preview.setWordWrap(True)
        lay_sc.addWidget(self.lbl_preview)
        lay_sc.addWidget(QLabel("<b>Scalebar detail</b> (enlarged, as exported)"))
        self.lbl_sb_detail = QLabel()
        self.lbl_sb_detail.setAlignment(Qt.AlignCenter)
        self.lbl_sb_detail.setMinimumHeight(70)
        self.lbl_sb_detail.setStyleSheet("background: white; border: 1px solid #999;")
        lay_sc.addWidget(self.lbl_sb_detail)
        self.lbl_sb_info = QLabel()
        self.lbl_sb_info.setWordWrap(True)
        lay_sc.addWidget(self.lbl_sb_info)
        side.addWidget(grp_sc, 1)

        self.cmb_bookmark.currentIndexChanged.connect(self._update_preview)
        self.cmb_round.currentIndexChanged.connect(self._scale_mode_changed)
        self.spn_scale.valueChanged.connect(self._update_preview)
        self.spn_w.valueChanged.connect(self._update_preview)
        self.spn_h.valueChanged.connect(self._update_preview)

        # 9-10. Export folder + file names
        grp_out = QGroupBox("9-10. Export (PNG @ %d dpi)" % DPI)
        form_out = QFormLayout(grp_out)
        proj_dir = QgsProject.instance().homePath() or os.path.expanduser("~")
        self.txt_folder = QLineEdit(os.path.join(proj_dir, "exports"))
        btn_browse = QPushButton("Browse...")
        btn_browse.clicked.connect(self._browse)
        row_folder = QHBoxLayout()
        row_folder.addWidget(self.txt_folder)
        row_folder.addWidget(btn_browse)
        self.txt_filename = QLineEdit(FILENAME_TEMPLATE)
        self.txt_filename.setToolTip("Fields: {theme} {bookmark} {layout} {scale} {page}")
        self.lbl_filename = QLabel()
        self.txt_filename.textChanged.connect(self._update_filename_example)
        self.lst_themes.itemChanged.connect(self._update_filename_example)
        self.cmb_bookmark.currentIndexChanged.connect(self._update_filename_example)
        form_out.addRow("Folder", row_folder)
        form_out.addRow("File name", self.txt_filename)
        form_out.addRow("", self.lbl_filename)
        root.addWidget(grp_out)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Create layout and export")
        buttons.accepted.connect(self._validate)
        buttons.rejected.connect(self.reject)
        outer.addWidget(buttons)

        if manual and manual[2] == "px":
            self.cmb_unit.setCurrentText("px")
        self._apply_preset(self.cmb_preset.currentText())
        self._scale_mode_changed()
        self._update_filename_example()
        self._fit_to_screen()

    def _fit_to_screen(self):
        """Open at the contents' natural size, capped to 90% of the available screen."""
        screen = None
        try:
            screen = (self.parent() or self).screen()
        except AttributeError:          # Qt < 5.14
            pass
        if screen is None:
            screen = QGuiApplication.primaryScreen()
        avail = screen.availableGeometry()
        hint = self._content.sizeHint()
        w = min(hint.width() + 40, int(avail.width() * 0.9))
        h = min(hint.height() + 80, int(avail.height() * 0.9))
        self.resize(w, h)
        self.move(avail.x() + (avail.width() - w) // 2, avail.y() + (avail.height() - h) // 2)

    def _bookmark_size_m(self):
        """Bookmark width/height on the ground in metres (project CRS), or None."""
        if not self.bookmarks:
            return None
        project = QgsProject.instance()
        ext = bookmark_extent_in_project_crs(project, self.bookmarks[self.cmb_bookmark.currentIndex()][1])
        f = QgsUnitTypes.fromUnitToUnitFactor(project.crs().mapUnits(), distance_unit("Meters"))
        return ext.width() * f, ext.height() * f

    def _scale_mode_changed(self, *_):
        self.spn_scale.setEnabled(self.cmb_round.currentData() == "manual")
        self._update_preview()

    def _update_preview(self, *_):
        """Recompute the scale (golden modes) and redraw the boundary on the page."""
        if not hasattr(self, "preview"):  # size widgets fire before the preview exists
            return
        page_mm = self.page_size()[:2]
        ground = self._bookmark_size_m()
        if not ground or min(ground) <= 0:
            self.preview.set_state(page_mm, None)
            self.lbl_preview.setText("No bookmark to preview.")
            return

        # scale at which the bookmark exactly fits the page (what zoomToExtent gives)
        fit = max(ground[0] * 1000 / page_mm[0], ground[1] * 1000 / page_mm[1])
        mode = self.cmb_round.currentData()
        if mode != "manual":
            self.spn_scale.blockSignals(True)
            self.spn_scale.setValue(int(round(golden_scale(fit, mode))))
            self.spn_scale.blockSignals(False)
        scale = self.spn_scale.value()

        boundary_mm = (ground[0] * 1000 / scale, ground[1] * 1000 / scale)
        self.preview.set_state(page_mm, boundary_mm)

        cover_w = boundary_mm[0] / page_mm[0] * 100
        cover_h = boundary_mm[1] / page_mm[1] * 100
        page_ground = (page_mm[0] * scale / 1e6, page_mm[1] * scale / 1e6)  # km
        text = ("Bookmark fits at 1:{:,.0f} &rarr; map at <b>1:{:,}</b><br>"
                "Boundary fills {:.0f}% of page width, {:.0f}% of height. "
                "Page covers {:.2f} &times; {:.2f} km on the ground."
                ).format(fit, scale, cover_w, cover_h, page_ground[0], page_ground[1])
        if cover_w > 100.5 or cover_h > 100.5:
            text += "<br><span style='color:#d33'><b>Boundary is cropped at this scale.</b></span>"
        self.lbl_preview.setText(text)
        if hasattr(self, "lbl_filename"):   # {scale} in the file name follows the scale
            self._update_filename_example()
        self._schedule_decorations()        # auto segments depend on the scale

    # ---- scalebar segments ----------------------------------------------
    def scalebar_options(self):
        fixed = self.cmb_seg_mode.currentData() == "fixed"
        return {"segments": self.spn_seg.value(), "segments_left": self.spn_seg_left.value(),
                "units": self.cmb_sb_units.currentData(),
                "segment_size": self.spn_seg_size.value() if fixed else None}

    def _segment_mode_changed(self, *_):
        fixed = self.cmb_seg_mode.currentData() == "fixed"
        self.spn_seg_size.setEnabled(fixed)
        if fixed and self._last_deco_info:  # start from what auto mode chose
            self.spn_seg_size.blockSignals(True)
            self.spn_seg_size.setValue(self._last_deco_info["per_segment"])
            self.spn_seg_size.blockSignals(False)
        self._schedule_decorations()

    def _schedule_decorations(self, *_):
        if hasattr(self, "_deco_timer"):
            self._deco_timer.start()

    def _render_decorations(self):
        """Rebuild the real scalebar + arrow at the current scale and show them."""
        if not hasattr(self, "lbl_sb_info"):
            return
        page_w, page_h = self.page_size()[:2]
        scale = self.spn_scale.value()
        extent = None
        if self.bookmarks:
            extent = bookmark_extent_in_project_crs(
                QgsProject.instance(), self.bookmarks[self.cmb_bookmark.currentIndex()][1])
        try:
            items, detail, info = decoration_preview(
                page_w, page_h, extent, scale, self.cmb_sb_corner.currentText(),
                self.cmb_sb_style.currentText(), self.cmb_na_corner.currentText(),
                self.chk_beside.isChecked(), self.scalebar_options())
        except Exception as e:  # keep the dialog usable if a combination fails to build
            self.preview.set_decorations(None)
            self.lbl_sb_detail.clear()
            self.lbl_sb_info.setText("<span style='color:#d33'>Scalebar preview failed: %s</span>" % e)
            return
        self._last_deco_info = info
        self.preview.set_decorations(items)

        target = QSize(max(100, self.lbl_sb_detail.width() - 12), 90)
        self.lbl_sb_detail.setPixmap(QPixmap.fromImage(detail).scaled(
            target, Qt.KeepAspectRatio, Qt.SmoothTransformation))

        self.spn_seg_size.setSuffix(" " + info["unit"])
        left = ""
        if info["segments_left"]:
            left = " + 1 left (split in %d)" % info["segments_left"]
        text = ("Scalebar: %d &times; %g %s%s = <b>%g %s</b>, bar %.1f mm long."
                % (info["segments"], info["per_segment"], info["unit"], left,
                   info["length"], info["unit"], info["bar_mm"]))
        usable = page_w - 2 * ITEM_MARGIN_MM
        if info["bar_mm"] > usable:
            text += ("<br><span style='color:#d33'><b>Bar is longer than the page "
                     "(%.0f mm available).</b></span>" % usable)
        elif info["bar_mm"] < 10:
            text += "<br><span style='color:#d33'><b>Bar is very short; use a larger segment size.</b></span>"
        self.lbl_sb_info.setText(text)

    def _set_all_themes(self, state):
        for i in range(self.lst_themes.count()):
            self.lst_themes.item(i).setCheckState(state)

    # Page size is held exactly as entered, self._exact = (w, h, "mm" | "px").
    # The boxes only display it in the chosen unit, so switching units or picking
    # a mm preset never rounds the page through whole pixels.

    def _apply_preset(self, text):
        size = self.presets.get(text)
        custom = size is None
        self.spn_w.setEnabled(custom)
        self.spn_h.setEnabled(custom)
        if size:
            self._exact = size
        self._display_exact()

    def _configure_size_boxes(self):
        """Ranges cover 100-30000 px in either unit; mm gets 2 decimals."""
        for s in (self.spn_w, self.spn_h):
            s.blockSignals(True)
            if self._unit == "mm":
                s.setDecimals(2)
                s.setRange(round(px_to_mm(100), 2), round(px_to_mm(30000), 2))
                s.setSingleStep(1.0)
            else:
                s.setDecimals(0)
                s.setRange(100, 30000)
                s.setSingleStep(10)
            s.setSuffix(" " + self._unit)
            s.blockSignals(False)

    def _change_unit(self, unit):
        """Show the width/height in another unit; the stored size is untouched."""
        self._unit = unit
        self._configure_size_boxes()
        self._display_exact()

    def _size_edited(self, *_):
        """User typed in the boxes: what they typed becomes the exact size."""
        self._exact = (self.spn_w.value(), self.spn_h.value(), self._unit)
        self._update_size_label()
        self._update_preview()

    def _display_exact(self):
        w, h, unit = self._exact
        if unit != self._unit:
            conv = px_to_mm if self._unit == "mm" else mm_to_px
            w, h = conv(w), conv(h)
        for box, v in ((self.spn_w, w), (self.spn_h, h)):
            box.blockSignals(True)
            box.setValue(v)
            box.blockSignals(False)
        self._update_size_label()
        self._update_preview()

    def page_size(self):
        """(page_w_mm, page_h_mm, px_w, px_h): mm exact for mm input, px exact for px input."""
        w, h, unit = self._exact
        if unit == "mm":
            return w, h, int(round(mm_to_px(w))), int(round(mm_to_px(h)))
        w, h = int(round(w)), int(round(h))
        return px_to_mm(w), px_to_mm(h), w, h

    def _update_size_label(self):
        w_mm, h_mm, w_px, h_px = self.page_size()
        self.lbl_size.setText("%g x %g mm  =  %d x %d px @ %d dpi"
                              % (round(w_mm, 3), round(h_mm, 3), w_px, h_px, DPI))

    def _browse(self):
        folder = QFileDialog.getExistingDirectory(self, "Select export folder", self.txt_folder.text())
        if folder:
            self.txt_folder.setText(folder)

    def _validate(self):
        if not self.bookmarks:
            QMessageBox.warning(self, "No bookmarks", "Create a spatial bookmark first (Ctrl+B).")
            return
        if not self.selected_themes():
            QMessageBox.warning(self, "No themes", "Select at least one map theme.")
            return
        if not self.txt_folder.text().strip():
            QMessageBox.warning(self, "No folder", "Select an export folder.")
            return
        if self._filename_preview() is None:
            QMessageBox.warning(self, "File name", "The file name pattern has an unknown {field}. "
                                "Use {theme} {bookmark} {layout} {scale} {page}.")
            return
        self.accept()

    def _filename_preview(self):
        """Example file names for the ticked themes, or None if the pattern is invalid."""
        bm = self.bookmarks[self.cmb_bookmark.currentIndex()][1].name() if self.bookmarks else ""
        fields = [{"theme": t, "bookmark": bm, "layout": self.txt_layout.text().strip() or LAYOUT_NAME,
                   "scale": "%d" % self.spn_scale.value(), "page": i + 1}
                  for i, t in enumerate(self.selected_themes())]
        try:
            return file_stems(self.txt_filename.text(), fields)
        except (KeyError, IndexError, ValueError):
            return None

    def _update_filename_example(self, *_):
        stems = self._filename_preview()
        if stems is None:
            self.lbl_filename.setText("<span style='color:#d33'>Unknown field in pattern</span>")
        elif stems:
            more = ", ..." if len(stems) > 2 else ""
            self.lbl_filename.setText("e.g. " + ", ".join(s + ".png" for s in stems[:2]) + more)
        else:
            self.lbl_filename.setText("")

    def selected_themes(self):
        return [self.lst_themes.item(i).text() for i in range(self.lst_themes.count())
                if self.lst_themes.item(i).checkState() == Qt.Checked]

    def settings(self):
        return {
            "bookmark": self.bookmarks[self.cmb_bookmark.currentIndex()][1],
            "themes": self.selected_themes(),
            "layout_name": self.txt_layout.text().strip() or LAYOUT_NAME,
            "page_mm": self.page_size()[:2],
            "px": self.page_size()[2:],
            "sb_corner": self.cmb_sb_corner.currentText(),
            "sb_style": self.cmb_sb_style.currentText(),
            "sb_opts": self.scalebar_options(),
            "na_corner": self.cmb_na_corner.currentText(),
            "na_beside_sb": self.chk_beside.isChecked(),
            "round_mode": self.cmb_round.currentData(),
            "manual_scale": self.spn_scale.value(),
            "folder": self.txt_folder.text().strip(),
            "filename_template": self.txt_filename.text().strip() or "{theme}",
        }


# --------------------------------------------------------------------------
# Layout building (steps 3-8)
# --------------------------------------------------------------------------
def create_layout(project, name, page_w, page_h, n_pages):
    """Step 3 + 7: fresh print layout with n pages of page_w x page_h mm."""
    manager = project.layoutManager()
    existing = manager.layoutByName(name)
    if existing:
        manager.removeLayout(existing)

    layout = QgsPrintLayout(project)
    layout.initializeDefaults()  # adds one A4 page
    layout.setName(name)
    pages = layout.pageCollection()
    pages.page(0).setPageSize(QgsLayoutSize(page_w, page_h, QgsUnitTypes.LayoutMillimeters))
    for _ in range(n_pages - 1):
        page = QgsLayoutItemPage(layout)
        page.setPageSize(QgsLayoutSize(page_w, page_h, QgsUnitTypes.LayoutMillimeters))
        pages.addPage(page)
    manager.addLayout(layout)
    return layout


def bookmark_extent_in_project_crs(project, bookmark):
    ext = bookmark.extent()
    if ext.crs().isValid() and ext.crs() != project.crs():
        xform = QgsCoordinateTransform(ext.crs(), project.crs(), project)
        return xform.transformBoundingBox(ext)
    return ext


def add_map(layout, page_index, page_w, page_h, extent, theme, round_mode, manual_scale=None):
    """Step 4 + 8: full-bleed map item following a theme.

    Scale is snapped to the golden series ('up' / 'nearest'), or set to
    manual_scale when round_mode is 'manual'. Either way the bookmark stays centred.
    """
    m = QgsLayoutItemMap(layout)
    m.setId("Map - %s" % theme)
    m.setFrameEnabled(False)
    m.setBackgroundEnabled(True)
    m.setBackgroundColor(QColor(255, 255, 255))
    layout.addLayoutItem(m)
    m.attemptResize(QgsLayoutSize(page_w, page_h, QgsUnitTypes.LayoutMillimeters))
    m.attemptMove(QgsLayoutPoint(0, 0, QgsUnitTypes.LayoutMillimeters), True, False, page_index)

    m.setFollowVisibilityPreset(True)
    m.setFollowVisibilityPresetName(theme)

    m.zoomToExtent(extent)            # fits bookmark, keeps page aspect ratio
    raw = m.scale()
    snapped = manual_scale if round_mode == "manual" else golden_scale(raw, round_mode)
    m.setScale(snapped)               # keeps the bookmark centre
    return m, raw, snapped


def add_scalebar(layout, map_item, page_index, page_w, page_h, corner, style, scale, opts=None):
    """Step 5: scalebar linked to the page's map.

    opts (see default_scalebar_options): segments, segments_left, units,
    segment_size (None = auto fit between SCALEBAR_MIN/MAX_WIDTH_MM).
    """
    opts = dict(default_scalebar_options(), **(opts or {}))
    sb = QgsLayoutItemScaleBar(layout)
    sb.setId("Scalebar - %s" % map_item.id())
    layout.addLayoutItem(sb)
    sb.setLinkedMap(map_item)
    sb.applyDefaultSize()             # first: it resets units/segments, overridden below
    sb.setStyle(style)

    if scalebar_unit(scale, opts["units"]) == "km":
        sb.setUnits(distance_unit("Kilometers"))
        sb.setUnitLabel("km")
    else:
        sb.setUnits(distance_unit("Meters"))
        sb.setUnitLabel("m")
    sb.setMapUnitsPerScaleBarUnit(1.0)  # label multiplier; units already converted

    sb.setNumberOfSegmentsLeft(int(opts["segments_left"]))
    sb.setNumberOfSegments(int(opts["segments"]))
    if opts["segment_size"]:
        sb.setSegmentSizeMode(QgsScaleBarSettings.SegmentSizeFixed)
        sb.setUnitsPerSegment(float(opts["segment_size"]))
    else:
        sb.setSegmentSizeMode(QgsScaleBarSettings.SegmentSizeFitWidth)
        sb.setMinimumBarWidth(SCALEBAR_MIN_WIDTH_MM)
        sb.setMaximumBarWidth(SCALEBAR_MAX_WIDTH_MM)
    sb.setHeight(SCALEBAR_HEIGHT_MM)

    fmt = QgsTextFormat()
    fmt.setFont(QFont(SCALEBAR_FONT))
    fmt.setSize(SCALEBAR_FONT_SIZE_PT)
    fmt.setSizeUnit(QgsUnitTypes.RenderPoints)
    fmt.setColor(QColor(0, 0, 0))
    buf = QgsTextBufferSettings()
    buf.setEnabled(SCALEBAR_TEXT_HALO)  # white halo keeps labels readable on the map
    buf.setSize(0.8)
    buf.setColor(QColor(255, 255, 255))
    fmt.setBuffer(buf)
    sb.setTextFormat(fmt)

    sb.setBackgroundEnabled(False)
    sb.setFrameEnabled(False)
    sb.setBoxContentSpace(0)            # item box hugs labels + bar, for arrow alignment

    sb.refreshItemSize()
    sb.refresh()
    place_item(sb, corner, page_index, page_w, page_h)
    return sb


def add_north_arrow(layout, map_item, scalebar, page_index, page_w, page_h, corner):
    """Step 6: north arrow (QGIS bundled SVG) linked to the page's map, sized to the scalebar."""
    pic = QgsLayoutItemPicture(layout)
    pic.setId("North arrow - %s" % map_item.id())
    layout.addLayoutItem(pic)
    svg_path = QgsSymbolLayerUtils.svgSymbolNameToPath(NORTH_ARROW_SVG, QgsPathResolver())
    pic.setPicturePath(svg_path)
    pic.setSvgFillColor(NORTH_ARROW_FILL)
    pic.setSvgStrokeColor(NORTH_ARROW_STROKE)
    pic.setSvgStrokeWidth(NORTH_ARROW_STROKE_WIDTH_MM)
    pic.setResizeMode(QgsLayoutItemPicture.Zoom)
    pic.setLinkedMap(map_item)
    pic.setNorthMode(QgsLayoutItemPicture.GridNorth)
    size_arrow_to_scalebar(pic, scalebar)
    pic.setBackgroundEnabled(False)
    pic.setFrameEnabled(False)
    place_item(pic, corner, page_index, page_w, page_h)
    return pic


def decoration_preview(page_w, page_h, extent, scale, corner, style, na_corner, beside,
                       sb_opts, dpi=254):
    """Real scalebar + north arrow for the dialog preview.

    Builds a throwaway one-page layout with the same functions the export uses,
    so segments, labels and positions match the output exactly. Returns
    (items, detail, info): items = [(QImage, QRectF page mm)] for the page
    overlay (transparent background), detail = QImage of the scalebar (with
    the arrow when it sits beside it) on white, info = segment details.
    """
    layout = QgsPrintLayout(QgsProject.instance())
    layout.initializeDefaults()
    page = layout.pageCollection().page(0)
    page.setPageSize(QgsLayoutSize(page_w, page_h, QgsUnitTypes.LayoutMillimeters))
    m = QgsLayoutItemMap(layout)
    layout.addLayoutItem(m)
    m.attemptResize(QgsLayoutSize(page_w, page_h, QgsUnitTypes.LayoutMillimeters))
    m.attemptMove(QgsLayoutPoint(0, 0, QgsUnitTypes.LayoutMillimeters), True, False, 0)
    m.zoomToExtent(extent if extent is not None and not extent.isEmpty()
                   else QgsRectangle(0, 0, 1000, 1000))
    m.setScale(scale)
    m.setVisibility(False)              # scalebar still measures it; only decorations render

    sb = add_scalebar(layout, m, 0, page_w, page_h, corner, style, scale, sb_opts)
    na = add_north_arrow(layout, m, sb, 0, page_w, page_h, na_corner)
    if beside:
        place_arrow_beside_scalebar(sb, na, corner, 0, page_w, page_h)

    exporter = QgsLayoutExporter(layout)
    page.setVisible(False)              # no white paper: overlay keeps the preview visible
    items = [(exporter.renderRegionToImage(it.sceneBoundingRect(), QSize(), dpi), it.sceneBoundingRect())
             for it in (sb, na)]
    page.setVisible(True)
    region = sb.sceneBoundingRect().united(na.sceneBoundingRect()) if beside else sb.sceneBoundingRect()
    if not beside:
        na.setVisibility(False)
    detail = exporter.renderRegionToImage(region.adjusted(-1, -1, 1, 1), QSize(), dpi)

    unit = sb.unitLabel()
    per_seg = sb.unitsPerSegment()
    n_right, n_left = sb.numberOfSegments(), sb.numberOfSegmentsLeft()
    length = per_seg * (n_right + (1 if n_left else 0))
    metres = length * (1000 if unit == "km" else 1)
    info = {"unit": unit, "per_segment": per_seg, "segments": n_right, "segments_left": n_left,
            "length": length, "bar_mm": metres * 1000 / scale}
    return items, detail, info


# --------------------------------------------------------------------------
# Export (step 9)
# --------------------------------------------------------------------------
def file_stems(template, fields_per_page):
    """File names (no extension) from FILENAME_TEMPLATE, made unique per run.

    Two themes that sanitise to the same name (e.g. "A/B" and "A_B") get
    "_2", "_3" ... instead of overwriting each other.
    """
    stems, used = [], set()
    for fields in fields_per_page:
        stem = safe_name(template.format(**fields))
        candidate, n = stem, 2
        while candidate.lower() in used:
            candidate, n = "%s_%d" % (stem, n), n + 1
        used.add(candidate.lower())
        stems.append(candidate)
    return stems


def export_pages(layout, stems, px, folder):
    """Render each page to its own PNG (named by stems) at exact pixel size and DPI."""
    os.makedirs(folder, exist_ok=True)
    exporter = QgsLayoutExporter(layout)
    dots_per_meter = int(round(DPI / 0.0254))
    written = []
    for i, stem in enumerate(stems):
        img = exporter.renderPageToImage(i, QSize(px[0], px[1]), DPI)
        img.setDotsPerMeterX(dots_per_meter)
        img.setDotsPerMeterY(dots_per_meter)
        path = os.path.join(folder, stem + ".png")
        if not img.save(path, "PNG"):
            raise IOError("Could not write %s" % path)
        written.append(path)
    return written


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def run():
    project = QgsProject.instance()
    if not project.mapThemeCollection().mapThemes():
        QMessageBox.warning(iface.mainWindow(), "No map themes",
                            "The project has no map themes. Create them in the Layers panel "
                            "(eye icon > Add Theme) and run again.")
        return

    dlg = LayoutAutomationDialog(iface.mainWindow())
    if not dlg.exec_():
        return
    s = dlg.settings()

    page_w, page_h = s["page_mm"]    # exact as entered; s["px"] only sizes the PNG
    themes = s["themes"]
    bookmark = s["bookmark"]
    extent = bookmark_extent_in_project_crs(project, bookmark)

    layout = create_layout(project, s["layout_name"], page_w, page_h, len(themes))

    report, name_fields = [], []
    for i, theme in enumerate(themes):
        m, raw, snapped = add_map(layout, i, page_w, page_h, extent, theme, s["round_mode"], s["manual_scale"])
        sb = add_scalebar(layout, m, i, page_w, page_h, s["sb_corner"], s["sb_style"], snapped, s["sb_opts"])
        na = add_north_arrow(layout, m, sb, i, page_w, page_h, s["na_corner"])
        if s["na_beside_sb"]:
            place_arrow_beside_scalebar(sb, na, s["sb_corner"], i, page_w, page_h)
        report.append("Page %d  %-25s 1:%s  (fit 1:%s)" % (i + 1, theme, f"{snapped:,.0f}", f"{raw:,.0f}"))
        name_fields.append({"theme": theme, "bookmark": bookmark.name(), "layout": s["layout_name"],
                            "scale": "%d" % round(snapped), "page": i + 1})

    layout.refresh()

    try:
        stems = file_stems(s["filename_template"], name_fields)
        files = export_pages(layout, stems, s["px"], s["folder"])
    except Exception as e:
        QMessageBox.critical(iface.mainWindow(), "Export failed", str(e))
        return

    print("\n".join(report))
    print("Exported %d image(s) to %s" % (len(files), s["folder"]))
    print("\n".join("  " + os.path.basename(f) for f in files))
    iface.messageBar().pushSuccess(
        "Layout automation",
        "Exported %d page(s) at %d dpi to %s" % (len(files), DPI, s["folder"]))
    iface.openLayoutDesigner(layout)


# Runs from the Python Console; the Urban Design Setup plugin imports this file as
# "qgis_layout_automation" and calls run() itself.
if __name__ != "qgis_layout_automation":
    run()
