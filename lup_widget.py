r"""
Export Land Use window: Land Use Plan layer -> Excel + PowerPoint with a live preview.

Runs two ways, both on QGIS's own Python/Qt:
  * inside QGIS, from the "Export Land Use" plugin button (uses layers in the project,
    follows selection and edits);
  * standalone, via Run_LUP_Widget.bat (no QGIS window needed).

Layout: settings on the left (scrollable), preview tabs on the right
(slides drawn from the same scene as the .pptx, plus the Excel tables),
log at the bottom. All panes are on splitters, so the window can be resized
and the panes dragged to any size.
"""

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
for _p in (HERE, os.path.join(HERE, "vendor")):
    if os.path.isdir(_p) and _p not in sys.path:
        # vendor (python-pptx) goes last so QGIS's own lxml/openpyxl win
        sys.path.append(_p) if _p.endswith("vendor") else sys.path.insert(0, _p)

import re

from qgis.PyQt.QtCore import Qt, QTimer, QRectF, QPointF, QSettings, QUrl, QEvent, QVariant
from qgis.PyQt.QtGui import (QColor, QPainter, QPen, QBrush, QFont, QFontDatabase,
                             QPainterPath, QDesktopServices, QKeySequence)
from qgis.PyQt.QtWidgets import (
    QApplication, QMainWindow, QWidget, QSplitter, QScrollArea, QTabWidget, QVBoxLayout,
    QHBoxLayout, QFormLayout, QGroupBox, QLabel, QPushButton, QCheckBox, QComboBox,
    QDoubleSpinBox, QLineEdit, QFileDialog, QPlainTextEdit, QTableWidget, QTableWidgetItem,
    QHeaderView, QToolButton, QSizePolicy, QMessageBox, QShortcut, QFrame, QColorDialog)
from qgis.core import (QgsProject, QgsVectorLayer, QgsMapLayerProxyModel, QgsDistanceArea,
                       QgsUnitTypes, QgsWkbTypes)
from qgis.gui import QgsMapLayerComboBox, QgsFieldComboBox

import lup_pipeline as L
import lup_colors as LC

NUMERIC = re.compile(r"^[\d.,%+\-/ ]+$")

EMU_PT = 12700
APP_NAME = "Export Land Use"


def _font_family():
    fams = set(QFontDatabase().families())
    for f in ("Aptos", "Segoe UI", "Arial"):
        if f in fams:
            return f
    return QFont().family()


def _py(v):
    if v is None or (isinstance(v, QVariant) and v.isNull()):
        return None
    if hasattr(v, "toString") and not isinstance(v, str):
        try:
            return v.toString("yyyy-MM-dd")
        except TypeError:
            return v.toString()
    return v


# --------------------------------------------------------------------------
# slide preview: paints the scene produced by lup_pipeline.build_slides()
# --------------------------------------------------------------------------
class SlideView(QWidget):
    BASE = 1000.0 / L.SLIDE_W  # px per EMU at 100%

    def __init__(self, parent=None):
        super().__init__(parent)
        self.scene = None
        self.zoom = 1.0
        self.family = _font_family()
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.on_ctrl_wheel = None

    def set_scene(self, scene):
        self.scene = scene
        self.update()

    def set_zoom(self, z):
        self.zoom = max(0.1, min(4.0, z))
        m = 24
        self.setFixedSize(int(L.SLIDE_W * self.s()) + 2 * m, int(L.SLIDE_H * self.s()) + 2 * m)
        self.update()

    def s(self):
        return self.BASE * self.zoom

    def wheelEvent(self, e):
        if e.modifiers() & Qt.ControlModifier and self.on_ctrl_wheel:
            self.on_ctrl_wheel(1.1 if e.angleDelta().y() > 0 else 1 / 1.1)
            e.accept()
        else:
            e.ignore()

    # -- painting helpers
    def font(self, size_pt, bold):
        f = QFont(self.family)
        f.setPixelSize(max(1, int(round(size_pt * EMU_PT * self.s()))))
        f.setBold(bool(bold))
        return f

    def draw_text(self, p, x, y, w, h, text, size, bold, color, align, anchor):
        s = self.s()
        lines = text.split("\n")
        fonts = [self.font(size, bold if i == 0 else False) for i in range(len(lines))]
        lh = size * EMU_PT * s * 1.2
        total = lh * len(lines)
        top = y * s if anchor == "t" else y * s + (h * s - total) / 2
        flag = {"l": Qt.AlignLeft, "c": Qt.AlignHCenter, "r": Qt.AlignRight}[align]
        p.setPen(QColor("#" + color))
        for i, (ln, f) in enumerate(zip(lines, fonts)):
            p.setFont(f)
            r = QRectF(x * s, top + i * lh, w * s, lh)
            p.drawText(r, flag | Qt.AlignVCenter | Qt.TextDontClip, ln)

    def draw_table(self, p, t):
        s = self.s()
        xs = [t["left"]]
        for w in t["widths"]:
            xs.append(xs[-1] + w)
        y = t["top"]
        for row in t["rows"]:
            h = row["h"]
            for j, c in enumerate(row["cells"]):
                if c is None:
                    continue
                x0, x1 = xs[j], xs[j + c["span"]]
                rect = QRectF(x0 * s, y * s, (x1 - x0) * s, h * s)
                if c["fill"]:
                    p.fillRect(rect, QColor("#" + c["fill"]))
                if c["bline"]:
                    p.setPen(QPen(QColor("#" + c["bline"]), max(1.0, 0.75 * EMU_PT * s)))
                    p.drawLine(QPointF(x0 * s, (y + h) * s), QPointF(x1 * s, (y + h) * s))
                if c["text"]:
                    ml, mr = 36000 + c["indent"], 36000
                    self.draw_text(p, x0 + ml, y, (x1 - x0) - ml - mr, h, c["text"], c["size"],
                                   c["bold"], c["color"], c["align"], "m")
            y += h

    def draw_chart(self, p, c):
        s = self.s()
        tot = float(sum(c["vals"])) or 1.0
        r = c["d"] / 2 * s
        cx, cy = c["cx"] * s, c["cy"] * s
        outer = QRectF(cx - r, cy - r, 2 * r, 2 * r)
        inner = None
        if c["kind"] == "doughnut":
            ri = r * (c["hole"] or 50) / 100
            inner = QPainterPath()
            inner.addEllipse(QRectF(cx - ri, cy - ri, 2 * ri, 2 * ri))
        cum = 0.0
        for v, col in zip(c["vals"], c["colors"]):
            frac = v / tot
            if col and frac > 0:
                path = QPainterPath(QPointF(cx, cy))
                path.arcTo(outer, 90 - cum * 360, -frac * 360)
                path.closeSubpath()
                if inner is not None:
                    path = path.subtracted(inner)
                edge = "BFBFBF" if col.upper() == "FFFFFF" else "FFFFFF"
                p.setPen(QPen(QColor("#" + edge), max(1.0, 0.75 * EMU_PT * s)))
                p.setBrush(QBrush(QColor("#" + col)))
                p.drawPath(path)
            cum += frac
        p.setBrush(Qt.NoBrush)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.TextAntialiasing)
        m = 24
        s = self.s()
        page = QRectF(m, m, L.SLIDE_W * s, L.SLIDE_H * s)
        p.fillRect(page.translated(4, 4), QColor(0, 0, 0, 40))
        p.fillRect(page, Qt.white)
        if not self.scene:
            p.setPen(QColor("#888"))
            p.drawText(page, Qt.AlignCenter, "No data - choose a land use layer")
            return
        p.translate(m, m)
        p.setClipRect(QRectF(0, 0, L.SLIDE_W * s, L.SLIDE_H * s))
        for it in self.scene["items"]:
            k = it["t"]
            if k == "text":
                self.draw_text(p, it["x"], it["y"], it["w"], it["h"], it["text"], it["size"],
                               it["bold"], it["color"], it["align"], it["anchor"])
            elif k == "table":
                self.draw_table(p, it)
            elif k == "chart":
                self.draw_chart(p, it)
            elif k == "line":
                p.setPen(QPen(QColor("#" + it["color"]), max(1.0, it["w"] * EMU_PT * s)))
                p.drawLine(QPointF(it["x0"] * s, it["y0"] * s), QPointF(it["x1"] * s, it["y1"] * s))
            elif k == "rect":
                pen = QPen(QColor("#" + it["color"]), max(1.0, it["w_pt"] * EMU_PT * s))
                if it["dash"]:
                    pen.setStyle(Qt.DashLine)
                p.setPen(pen)
                p.setBrush(Qt.NoBrush)
                p.drawRect(QRectF(it["x"] * s, it["y"] * s, it["w"] * s, it["h"] * s))
        p.end()


class SlidePane(QWidget):
    """Scrollable slide preview with zoom controls (Fit / % / Ctrl+wheel)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.view = SlideView()
        self.view.on_ctrl_wheel = lambda f: self.set_zoom(self.view.zoom * f)
        self.area = QScrollArea()
        self.area.setWidget(self.view)
        self.area.setAlignment(Qt.AlignCenter)
        self.area.setStyleSheet("QScrollArea { background: #e9e9ec; }")
        self.area.viewport().installEventFilter(self)
        self.zoom_box = QComboBox()
        self.zoom_box.setEditable(True)  # type any %, e.g. 130%
        self.zoom_box.setMinimumContentsLength(8)
        self.zoom_box.addItems(["Fit", "Fit width", "50%", "75%", "100%", "125%", "150%", "200%"])
        self.zoom_box.setCurrentText("Fit")
        self.zoom_box.currentTextChanged.connect(self.apply_zoom)
        bar = QHBoxLayout()
        bar.setContentsMargins(4, 4, 4, 0)
        for txt, f in (("−", 1 / 1.25), ("+", 1.25)):
            b = QToolButton()
            b.setText(txt)
            b.setToolTip("Zoom out" if f < 1 else "Zoom in")
            b.clicked.connect(lambda _, f=f: self.set_zoom(self.view.zoom * f))
            bar.addWidget(b)
        bar.addWidget(QLabel("Zoom"))
        bar.addWidget(self.zoom_box)
        bar.addStretch(1)
        bar.addWidget(QLabel("Ctrl + mouse wheel to zoom"))
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addLayout(bar)
        lay.addWidget(self.area, 1)
        self.view.set_zoom(1.0)

    def eventFilter(self, obj, ev):
        if obj is self.area.viewport() and ev.type() == QEvent.Resize:
            if self.zoom_box.currentText().startswith("Fit"):
                QTimer.singleShot(0, lambda: self.apply_zoom(self.zoom_box.currentText()))
        return False

    def fit_zoom(self, width_only):
        vp = self.area.viewport().size()
        zw = (vp.width() - 52) / (L.SLIDE_W * SlideView.BASE)
        zh = (vp.height() - 52) / (L.SLIDE_H * SlideView.BASE)
        return zw if width_only else min(zw, zh)

    def apply_zoom(self, txt):
        if txt == "Fit":
            self.view.set_zoom(self.fit_zoom(False))
        elif txt == "Fit width":
            self.view.set_zoom(self.fit_zoom(True))
        else:
            try:
                z = float(txt.strip().rstrip("%").replace(",", ".")) / 100
            except ValueError:
                return
            if z >= 0.1:
                self.view.set_zoom(z)

    def set_zoom(self, z):
        self.view.set_zoom(z)
        self.zoom_box.blockSignals(True)
        self.zoom_box.setEditText(f"{int(round(self.view.zoom * 100))}%")
        self.zoom_box.blockSignals(False)


# --------------------------------------------------------------------------
# table previews (values of the Excel sheets)
# --------------------------------------------------------------------------
def make_table():
    t = QTableWidget()
    t.setEditTriggers(QTableWidget.NoEditTriggers)
    t.setAlternatingRowColors(False)
    t.verticalHeader().setVisible(False)
    t.horizontalHeader().setStretchLastSection(False)
    t.setSelectionBehavior(QTableWidget.SelectRows)
    t.setWordWrap(False)
    t.verticalHeader().setDefaultSectionSize(t.fontMetrics().height() + 8)
    return t


def fill_table(t, headers, rows):
    """rows: list of (cells, style) where style = dict(fill=hex for col 0, bold=bool, rowfill=hex)."""
    t.clear()
    t.setColumnCount(len(headers))
    t.setRowCount(len(rows))
    t.setHorizontalHeaderLabels(headers)
    for i, (cells, st) in enumerate(rows):
        for j, v in enumerate(cells):
            it = QTableWidgetItem("" if v is None else str(v))
            if j > 0 and NUMERIC.match(it.text() or "x"):
                it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            fill = st.get("rowfill") or (st.get("fill") if j == 0 else None)
            if fill:
                it.setBackground(QColor("#" + fill))
                it.setForeground(QColor("#FFFFFF" if L.is_dark(fill) else "#000000"))
            if st.get("bold"):
                f = it.font()
                f.setBold(True)
                it.setFont(f)
            t.setItem(i, j, it)
    t.resizeColumnsToContents()
    for j in range(len(headers)):
        t.setColumnWidth(j, max(70, min(320, t.columnWidth(j) + 16)))


# --------------------------------------------------------------------------
# main window
# --------------------------------------------------------------------------
class LupWindow(QMainWindow):
    FIELD_KEYS = [("F_TYPE", "Saleable / Non-Saleable", "Type", False),
                  ("F_SUBZ", "Land use (SUBZONA)", "SUBZONA", False),
                  ("F_AREA", "Area (sqm)", "Area_sqm", True),
                  ("F_INDT", "Industrial type", "IND_TYPE", True),
                  ("F_PHASE", "Phase", "Phase", True),
                  ("F_PARCEL", "Parcel ID", "Parcel_ID", True),
                  ("F_POP", "Population", "Pop_Total", True),
                  ("F_POPTYPE", "Population type", "Pop_Type", True)]

    def __init__(self, iface=None, parent=None, lib_dir=HERE):
        super().__init__(parent)
        self.iface = iface
        self.lib_dir = lib_dir
        self.settings = QSettings("LUP", "LupTables")
        self.cfg = LC.load_config(lib_dir)
        self.lib = []
        self.colors = None
        self.data = None          # (fields, rows, S)
        self.layer = None
        self.last_outputs = (None, None)
        self.setWindowTitle(APP_NAME + " - Land Use Plan to Excel / PowerPoint")
        self.setMinimumSize(640, 420)
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.setInterval(350)
        self.timer.timeout.connect(self.refresh)
        self.build_ui()
        self.restore()
        QShortcut(QKeySequence("F5"), self, activated=self.refresh)
        QShortcut(QKeySequence("Ctrl+S"), self, activated=self.generate)
        QTimer.singleShot(0, self.refresh)

    # ---- UI ---------------------------------------------------------------
    def build_ui(self):
        form_host = QWidget()
        side = QVBoxLayout(form_host)

        g_in = QGroupBox("1. Land use layer")
        f = QFormLayout(g_in)
        self.layer_cb = QgsMapLayerComboBox()
        self.layer_cb.setFilters(QgsMapLayerProxyModel.PolygonLayer)
        self.layer_cb.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.layer_cb.setMinimumContentsLength(8)
        self.layer_cb.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.layer_cb.layerChanged.connect(self.on_layer_changed)
        browse = QPushButton("Open file…")
        browse.setToolTip("Add a shapefile / GeoPackage layer")
        browse.clicked.connect(self.browse_layer)
        row = QHBoxLayout()
        row.addWidget(self.layer_cb, 1)
        row.addWidget(browse)
        f.addRow(row)
        self.sel_only = QCheckBox("Selected features only")
        self.sel_only.setToolTip("Report on the selection only; its own area is the grand total.")
        self.sel_only.toggled.connect(self.schedule)
        self.sel_only.toggled.connect(lambda on: self.site_total.setEnabled(not on))
        f.addRow(self.sel_only)
        self.auto = QCheckBox("Live preview (follow edits && selection)")
        self.auto.setChecked(True)
        f.addRow(self.auto)
        self.count_lbl = QLabel("")
        self.count_lbl.setStyleSheet("color:#666")
        f.addRow(self.count_lbl)
        side.addWidget(g_in)

        g_set = QGroupBox("2. Settings")
        f = QFormLayout(g_set)
        self.site_total = QDoubleSpinBox()
        self.site_total.setRange(0, 1e12)
        self.site_total.setDecimals(0)
        self.site_total.setSingleStep(1000)
        self.site_total.setGroupSeparatorShown(True)
        self.site_total.setValue(5_000_000)
        self.site_total.setToolTip("GRAND TOTAL in sqm. Road absorbs the difference to the data.\n0 = sum of the layer.")
        self.site_total.valueChanged.connect(self.schedule)
        f.addRow("Site total (sqm)", self.site_total)
        self.area_src = QComboBox()
        self.area_src.addItems(["Area field", "Polygon geometry"])
        self.area_src.currentIndexChanged.connect(self.schedule)
        f.addRow("Parcel area from", self.area_src)
        self.style = QComboBox()
        self.style.addItems(["Indonesian 1.234,56", "English 1,234.56"])
        self.style.currentIndexChanged.connect(self.schedule)
        f.addRow("Number format", self.style)
        self.template = QLineEdit()
        tb = QToolButton()
        tb.setText("…")
        tb.clicked.connect(lambda: self.pick_file(self.template, "PowerPoint (*.pptx)"))
        row = QHBoxLayout()
        row.addWidget(self.template, 1)
        row.addWidget(tb)
        f.addRow("PPT template", row)
        side.addWidget(g_set)

        g_col = QGroupBox("Colour coding")
        g_col.setToolTip("One colour scheme for the layer symbology, Excel cells and slides")
        f = QFormLayout(g_col)
        self.color_field = QgsFieldComboBox()
        self.color_field.setToolTip("Field whose values get the colours (Excel conditional formatting, slides, layer style)")
        self.color_field.fieldChanged.connect(self.on_color_field)
        f.addRow("Colour by field", self.color_field)
        self.priority = QComboBox()
        for key, (label, _) in LC.PRIORITIES.items():
            self.priority.addItem(label, key)
        self.priority.currentIndexChanged.connect(self.on_priority)
        f.addRow("Colour source", self.priority)
        self.library = QLineEdit()
        self.library.setToolTip("RDTR / ATR-BPN colour code workbook (Zona | Kode Z | Sub-Zona | Kode SZ | RGB)")
        self.library.editingFinished.connect(self.on_library)
        tb = QToolButton()
        tb.setText("…")
        tb.clicked.connect(self.browse_library)
        row = QHBoxLayout()
        row.addWidget(self.library, 1)
        row.addWidget(tb)
        f.addRow("Colour library", row)
        self.lib_info = QLabel("")
        self.lib_info.setStyleSheet("color:#666")
        self.lib_info.setWordWrap(True)
        f.addRow(self.lib_info)
        row = QHBoxLayout()
        b1 = QPushButton("Use the layer's style field")
        b1.setToolTip("Colour by the field the layer's categorised symbology uses")
        b1.clicked.connect(self.use_layer_style_field)
        b2 = QPushButton("Apply these colours to the layer style")
        b2.setToolTip("Write these colours to the layer as categorised symbology")
        b2.clicked.connect(self.apply_layer_style)
        b3 = QPushButton("Clear manual colours")
        b3.clicked.connect(self.clear_overrides)
        for b in (b1, b2, b3):  # stacked so they fit a narrow panel
            f.addRow(b)
        side.addWidget(g_col)

        self.g_fields = QGroupBox("Field names (click to show)")
        self.g_fields.setCheckable(True)
        self.g_fields.setChecked(False)
        fields_host = QWidget()
        f = QFormLayout(fields_host)
        f.setContentsMargins(0, 0, 0, 0)
        self.field_cbs = {}
        for key, label, default, optional in self.FIELD_KEYS:
            cb = QgsFieldComboBox()
            cb.setAllowEmptyFieldName(optional)
            cb.fieldChanged.connect(self.schedule)
            self.field_cbs[key] = (cb, default)
            f.addRow(label, cb)
        lay = QVBoxLayout(self.g_fields)
        lay.addWidget(fields_host)
        fields_host.setVisible(False)
        self.g_fields.toggled.connect(fields_host.setVisible)
        side.addWidget(self.g_fields)

        g_out = QGroupBox("3. Output")
        f = QFormLayout(g_out)
        self.out_dir = QLineEdit()
        tb = QToolButton()
        tb.setText("…")
        tb.clicked.connect(self.pick_dir)
        row = QHBoxLayout()
        row.addWidget(self.out_dir, 1)
        row.addWidget(tb)
        f.addRow("Folder", row)
        self.base_name = QLineEdit()
        self.base_name.setPlaceholderText("layer name")
        f.addRow("File name", self.base_name)
        self.want_xlsx = QCheckBox("Excel database (.xlsx)")
        self.want_pptx = QCheckBox("PowerPoint tables (.pptx)")
        self.want_xlsx.setChecked(True)
        self.want_pptx.setChecked(True)
        f.addRow(self.want_xlsx)
        self.want_pivots = QCheckBox("Summary sheets as pivot tables (uses Excel)")
        self.want_pivots.setToolTip("Pivot tables on the Source table, built by Microsoft Excel.\n"
                                    "Unticked or without Excel: SUMIFS formula sheets.")
        self.want_pivots.setChecked(True)
        self.want_xlsx.toggled.connect(self.want_pivots.setEnabled)
        f.addRow(self.want_pivots)
        f.addRow(self.want_pptx)
        self.open_after = QCheckBox("Open files when done")
        self.open_after.setChecked(True)
        f.addRow(self.open_after)
        side.addWidget(g_out)

        self.gen_btn = QPushButton("Generate files  (Ctrl+S)")
        self.gen_btn.setMinimumHeight(34)
        self.gen_btn.setStyleSheet("font-weight:bold")
        self.gen_btn.clicked.connect(self.generate)
        side.addWidget(self.gen_btn)
        row = QHBoxLayout()
        self.open_x = QPushButton("Open Excel")
        self.open_p = QPushButton("Open PPT")
        self.open_d = QPushButton("Open folder")
        self.open_x.clicked.connect(lambda: self.open_path(self.last_outputs[0]))
        self.open_p.clicked.connect(lambda: self.open_path(self.last_outputs[1]))
        self.open_d.clicked.connect(lambda: self.open_path(self.out_dir.text()))
        for b in (self.open_x, self.open_p, self.open_d):
            row.addWidget(b)
        side.addLayout(row)
        ref = QPushButton("Refresh preview  (F5)")
        ref.clicked.connect(self.refresh)
        side.addWidget(ref)
        side.addStretch(1)

        left = QScrollArea()
        left.setWidgetResizable(True)
        left.setWidget(form_host)
        left.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)  # fields shrink, panel scrolls vertically
        left.setMinimumWidth(280)
        left.setFrameShape(QFrame.NoFrame)

        self.tabs = QTabWidget()
        self.slide_panes = [SlidePane(), SlidePane()]
        self.tabs.addTab(self.slide_panes[0], "Slide 1 · Summary")
        self.tabs.addTab(self.slide_panes[1], "Slide 2 · Industrial")
        self.tabs.addTab(self.build_colour_tab(), "Colours")
        self.tables = {}
        for name in ("Overall MP", "Industrial", "Population", "Phasing", "Intensitas", "Checks"):
            t = make_table()
            self.tables[name] = t
            self.tabs.addTab(t, name)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(2000)

        right = QSplitter(Qt.Vertical)
        right.addWidget(self.tabs)
        right.addWidget(self.log)
        right.setStretchFactor(0, 5)
        right.setStretchFactor(1, 1)
        right.setSizes([600, 110])
        self.vsplit = right

        self.hsplit = QSplitter(Qt.Horizontal)
        self.hsplit.addWidget(left)
        self.hsplit.addWidget(right)
        self.hsplit.setStretchFactor(1, 1)
        self.hsplit.setSizes([370, 900])
        self.setCentralWidget(self.hsplit)
        self.statusBar().showMessage("Ready")

    # ---- settings -----------------------------------------------------------
    def restore(self):
        st = self.settings
        g = st.value("geometry")
        if g is not None:
            self.restoreGeometry(g)
        else:
            self.resize(1280, 820)
        for sp, key in ((self.hsplit, "hsplit"), (self.vsplit, "vsplit")):
            v = st.value(key)
            if v is not None:
                sp.restoreState(v)
        self.site_total.setValue(float(st.value("site_total", 5_000_000)))
        self.area_src.setCurrentIndex(int(st.value("area_src", 0)))
        self.style.setCurrentIndex(int(st.value("style", 0)))
        tpl = st.value("template", "")
        if not tpl or not os.path.exists(tpl):
            for c in (os.path.join(self.lib_dir, "LUP_template.pptx"),
                      os.path.join(HERE, "LUP_template.pptx")):
                if os.path.exists(c):
                    tpl = c
                    break
        self.template.setText(tpl or "")
        self.out_dir.setText(st.value("out_dir", "") or os.path.join(HERE, "Output"))
        self.priority.blockSignals(True)
        self.priority.setCurrentIndex(max(0, self.priority.findData(self.cfg.get("priority", "layer"))))
        self.priority.blockSignals(False)
        self.library.setText(LC.library_path(self.cfg, self.lib_dir) or self.cfg.get("library", ""))
        self.load_library()
        if self.layer_cb.count() == 0:  # standalone: reopen last file / sample shapefile
            last = st.value("last_file", "")
            if not last or not os.path.exists(last):
                shp_dir = os.path.join(HERE, "Shapefile")
                shp = [f for f in os.listdir(shp_dir) if f.lower().endswith(".shp")] if os.path.isdir(shp_dir) else []
                last = os.path.join(shp_dir, shp[0]) if shp else ""
            if last:
                self.add_file_layer(last)
        self.on_layer_changed(self.layer_cb.currentLayer())

    def closeEvent(self, e):
        st = self.settings
        st.setValue("geometry", self.saveGeometry())
        st.setValue("hsplit", self.hsplit.saveState())
        st.setValue("vsplit", self.vsplit.saveState())
        st.setValue("site_total", self.site_total.value())
        st.setValue("area_src", self.area_src.currentIndex())
        st.setValue("style", self.style.currentIndex())
        st.setValue("template", self.template.text())
        st.setValue("out_dir", self.out_dir.text())
        self.watch(None)
        super().closeEvent(e)

    # ---- layer handling ---------------------------------------------------
    def browse_layer(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open land use layer", self.settings.value("last_file", ""),
                                              "Vector files (*.shp *.gpkg *.geojson *.json *.kml *.tab);;All files (*)")
        if path:
            self.add_file_layer(path)

    def add_file_layer(self, path):
        lyr = QgsVectorLayer(path, os.path.splitext(os.path.basename(path))[0], "ogr")
        if not lyr.isValid():
            self.say(f"Could not open {path}")
            return
        if QgsWkbTypes.geometryType(lyr.wkbType()) != QgsWkbTypes.PolygonGeometry:
            self.say(f"{path} is not a polygon layer")
            return
        QgsProject.instance().addMapLayer(lyr)
        self.layer_cb.setLayer(lyr)
        self.settings.setValue("last_file", path)

    def watch(self, layer):
        sigs = ("selectionChanged", "attributeValueChanged", "geometryChanged", "featureAdded",
                "featuresDeleted", "afterCommitChanges", "afterRollBack", "rendererChanged", "styleChanged")
        if self.layer is not None:
            for s in sigs:
                try:
                    getattr(self.layer, s).disconnect(self.on_layer_edit)
                except (TypeError, RuntimeError):
                    pass
        self.layer = layer
        if layer is not None:
            for s in sigs:
                getattr(layer, s).connect(self.on_layer_edit)

    def on_layer_edit(self, *args):
        if self.auto.isChecked():
            self.schedule()

    def on_layer_changed(self, layer):
        self.watch(layer)
        for key, (cb, default) in self.field_cbs.items():
            cb.blockSignals(True)
            cb.setLayer(layer)
            if layer is not None and layer.fields().indexFromName(default) >= 0:
                cb.setField(default)
            elif cb.allowEmptyFieldName():
                cb.setField("")
            cb.blockSignals(False)
        if layer is not None:
            self.base_name.setPlaceholderText(layer.name())
        # colour field: the layer's categorised style field, else the saved one, else SUBZONA
        self.color_field.blockSignals(True)
        self.color_field.setLayer(layer)
        if layer is not None:
            for f in (LC.renderer_field(layer), self.cfg.get("field"), L.F_SUBZ):
                if f and layer.fields().indexFromName(f) >= 0:
                    self.color_field.setField(f)
                    break
        self.color_field.blockSignals(False)
        self.schedule()

    # ---- colour settings ----------------------------------------------------
    def load_library(self):
        path = self.library.text().strip()
        self.lib = LC.load_library(path) if path and os.path.exists(path) else []
        if self.lib:
            self.lib_info.setText(f"{len(self.lib)} colours from {os.path.basename(path)}")
        else:
            self.lib_info.setText("Colour library not found - choose the RDTR colour code workbook")

    def browse_library(self):
        p, _ = QFileDialog.getOpenFileName(self, "Colour library (RDTR colour code)", self.library.text(),
                                           "Excel (*.xlsx *.xlsm)")
        if p:
            self.library.setText(p)
            self.on_library()

    def on_library(self):
        self.cfg["library"] = self.library.text().strip()
        LC.save_config(self.cfg, self.lib_dir)
        self.load_library()
        self.schedule()

    def on_priority(self, *_):
        self.cfg["priority"] = self.priority.currentData()
        LC.save_config(self.cfg, self.lib_dir)
        self.schedule()

    def on_color_field(self, f):
        if f:
            self.cfg["field"] = f
            LC.save_config(self.cfg, self.lib_dir)
        self.schedule()

    def use_layer_style_field(self):
        lyr = self.layer_cb.currentLayer()
        f = LC.renderer_field(lyr) if lyr else None
        if f:
            self.color_field.setField(f)
            self.statusBar().showMessage(f"Colouring by '{f}' (the layer's categorised style field)")
        else:
            QMessageBox.information(self, APP_NAME, "The layer has no categorised symbology.\n"
                                    "Pick the field yourself, or use 'Apply to layer style' to create one.")

    def apply_layer_style(self):
        lyr = self.layer_cb.currentLayer()
        if lyr is None or not self.colors:
            return
        where = ("Outline and other symbol settings are kept. Save the QGIS project to keep it."
                 if self.iface is not None else
                 f"It is saved as the layer's default style ({os.path.splitext(lyr.source().split('|')[0])[0]}.qml),\n"
                 "so QGIS shows these colours whenever the layer is opened.")
        if QMessageBox.question(self, APP_NAME,
                                f"Replace the symbology of '{lyr.name()}' with a categorised style on "
                                f"'{self.colors['field']}' using these {len(self.colors['scheme'])} colours?\n\n"
                                + where) != QMessageBox.Yes:
            return
        LC.apply_to_layer(lyr, self.colors["field"], self.colors["scheme"])
        if self.iface is not None:
            try:
                self.iface.layerTreeView().refreshLayerSymbology(lyr.id())
            except Exception:
                pass
        else:
            try:
                msg, ok = lyr.saveDefaultStyle()
            except TypeError:
                msg, ok = lyr.saveDefaultStyle(lyr.StyleCategory.AllStyleCategories)
            self.say(("Saved default style: " if ok else "Could not save default style: ") + str(msg))
        self.say(f"Layer style updated: {len(self.colors['scheme'])} categories on '{self.colors['field']}'")
        self.schedule()

    def clear_overrides(self):
        f = self.color_field.currentField()
        if (self.cfg.get("overrides") or {}).pop(f, None) is not None:
            LC.save_config(self.cfg, self.lib_dir)
        self.schedule()

    def set_alias(self, value, code):
        al = self.cfg.setdefault("aliases", {})
        if code == "__auto__":
            al.pop(value, None)
        else:
            al[value] = "-" if code == "__none__" else code
        LC.save_config(self.cfg, self.lib_dir)
        self.schedule()

    def pick_colour(self, value, current):
        c = QColorDialog.getColor(QColor("#" + current), self, f"Colour for '{value}'")
        if c.isValid():
            self.cfg.setdefault("overrides", {}).setdefault(self.colors["field"], {})[value] = c.name()[1:].upper()
            LC.save_config(self.cfg, self.lib_dir)
            self.schedule()

    def build_colour_tab(self):
        host = QWidget()
        lay = QVBoxLayout(host)
        self.col_info = QLabel("")
        self.col_info.setWordWrap(True)
        lay.addWidget(self.col_info)
        sp = QSplitter(Qt.Vertical)
        self.color_table = make_table()
        self.color_table.setSelectionMode(QTableWidget.NoSelection)
        self.group_table = make_table()
        self.group_table.setSelectionMode(QTableWidget.NoSelection)
        w2 = QWidget()
        l2 = QVBoxLayout(w2)
        l2.setContentsMargins(0, 0, 0, 0)
        self.group_lbl = QLabel("")
        l2.addWidget(self.group_lbl)
        l2.addWidget(self.group_table)
        sp.addWidget(self.color_table)
        sp.addWidget(w2)
        sp.setSizes([500, 200])
        lay.addWidget(sp, 1)
        return host

    def fill_colour_tab(self, rows):
        c = self.colors
        field = c["field"]
        cnt, area = {}, {}
        for r in rows:
            v = r.get(field)
            if v not in (None, ""):
                k = str(v).strip()
                cnt[k] = cnt.get(k, 0) + 1
                area[k] = area.get(k, 0.0) + L.num(r.get(L.F_AREA))
        lyr = self.layer_cb.currentLayer()
        rf = LC.renderer_field(lyr) if lyr else None
        sync = ("in sync with the layer's symbology" if rf == field else
                f"layer is styled on '{rf}', not '{field}'" if rf else "layer has no categorised symbology")
        n_lib = sum(1 for i in c["info"].values() if i["lib_entry"])
        self.col_info.setText(
            f"<b>{len(c['info'])}</b> values of <b>{field}</b> · {n_lib} matched in the colour library · {sync}.<br>"
            "Pick a library entry per value, or click a colour to set it by hand. The <i>Excel cell / slide</i> "
            "column shows exactly how the value is formatted in the workbook and on the slides.")
        hdr = ["Value", "Features", "Area (ha)", "Layer symbol", "Colour library match", "Library",
               "Used colour", "Source", "Excel cell / slide"]
        t = self.color_table
        t.clear()
        t.setColumnCount(len(hdr))
        t.setHorizontalHeaderLabels(hdr)
        t.setRowCount(len(c["info"]))
        lib_items = [(e["label"], e["kode_sz"] or e["kode_z"], e["hex"]) for e in self.lib]
        user_alias = {LC.norm(k): v for k, v in (self.cfg.get("aliases") or {}).items()}

        def swatch(hx, text=None):
            it = QTableWidgetItem(text if text is not None else ("#" + hx if hx else "—"))
            if hx:
                it.setBackground(QColor("#" + hx))
                it.setForeground(QColor("#FFFFFF" if LC.is_dark(hx) else "#000000"))
            return it

        for i, (v, inf) in enumerate(c["info"].items()):
            t.setItem(i, 0, QTableWidgetItem(v))
            t.setItem(i, 1, QTableWidgetItem(str(cnt.get(v, 0))))
            t.setItem(i, 2, QTableWidgetItem(L.fmt_dec(area.get(v, 0.0) / 1e4)))
            for j in (1, 2):
                t.item(i, j).setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            t.setItem(i, 3, swatch(inf["layer_hex"]))
            cb = QComboBox()
            e = inf["lib_entry"]
            cb.addItem(f"(auto) {e['label']}" if e and LC.norm(v) not in user_alias else "(auto)", "__auto__")
            cb.addItem("(no library colour)", "__none__")
            for label, code, hx in lib_items:
                cb.addItem(label, code)
            chosen = user_alias.get(LC.norm(v))
            if chosen == "-":
                cb.setCurrentIndex(1)
            elif chosen:
                k = cb.findData(chosen)
                cb.setCurrentIndex(k if k >= 0 else 0)
            cb.currentIndexChanged.connect(lambda k, v=v, cb=cb: self.set_alias(v, cb.itemData(k)))
            t.setCellWidget(i, 4, cb)
            t.setItem(i, 5, swatch(e["hex"] if e else None))
            btn = QPushButton("#" + inf["hex"])
            fg = "#FFFFFF" if LC.is_dark(inf["hex"]) else "#000000"
            btn.setStyleSheet(f"background:#{inf['hex']}; color:{fg}; border:1px solid #888; padding:2px 6px;")
            btn.setToolTip("Click to choose a colour by hand")
            btn.clicked.connect(lambda _, v=v, hx=inf["hex"]: self.pick_colour(v, hx))
            t.setCellWidget(i, 6, btn)
            t.setItem(i, 7, QTableWidgetItem(LC.source_text(inf)))
            t.setItem(i, 8, swatch(inf["hex"], v))
        t.resizeColumnsToContents()
        t.setColumnWidth(4, max(260, t.columnWidth(4)))
        t.setColumnWidth(8, max(160, t.columnWidth(8)))

        g = self.group_table
        if field == L.F_SUBZ:
            self.group_lbl.setText(f"Slides and pivot rows use the {field} colours above directly.")
            g.setVisible(False)
            return
        g.setVisible(True)
        self.group_lbl.setText(f"<b>{L.F_SUBZ}</b> rows on the slides and pivots take the colour of their main "
                               f"<b>{field}</b> value (by area):")
        rows_ = [([k, "#" + hx], {"fill": hx}) for k, hx in sorted(c["groups"].items())]
        fill_table(g, [L.F_SUBZ, "Colour"], rows_)

    # ---- compute ------------------------------------------------------------
    def schedule(self, *args):
        self.timer.start()

    def say(self, msg):
        self.log.appendPlainText(msg)

    def effective_site_total(self):
        # a selection is a part of the site: report it on its own total, no Road plug
        return 0 if self.sel_only.isChecked() else self.site_total.value()

    def apply_fields(self):
        for key, (cb, default) in self.field_cbs.items():
            setattr(L, key, cb.currentField() or default)

    def read_layer(self):
        lyr = self.layer_cb.currentLayer()
        if lyr is None:
            return None
        self.apply_fields()
        use_geom = self.area_src.currentIndex() == 1
        names = [f.name() for f in lyr.fields()]
        fields = list(names)
        if L.F_AREA not in fields:
            if not use_geom:
                raise ValueError(f"Area field '{L.F_AREA}' not found - choose another field or 'Polygon geometry'.")
            fields.append(L.F_AREA)
        da = QgsDistanceArea()
        da.setSourceCrs(lyr.crs(), QgsProject.instance().transformContext())
        if lyr.crs().isGeographic():
            da.setEllipsoid(QgsProject.instance().ellipsoid() or "EPSG:7030")
        feats = lyr.getSelectedFeatures() if self.sel_only.isChecked() else lyr.getFeatures()
        rows = []
        for ft in feats:
            rec = {n: _py(v) for n, v in zip(names, ft.attributes())}
            g = ft.geometry()
            a = da.convertAreaMeasurement(da.measureArea(g), QgsUnitTypes.AreaSquareMeters) if g and not g.isEmpty() else 0.0
            rec["_geom_area"] = a
            if use_geom:
                rec[L.F_AREA] = round(a, 2)
            rows.append(rec)
        return fields, rows

    def refresh(self):
        try:
            L.NUMBER_STYLE = ["id", "en"][self.style.currentIndex()]
            res = self.read_layer()
            if not res or not res[1]:
                self.data = None
                for p in self.slide_panes:
                    p.view.set_scene(None)
                self.count_lbl.setText("No features" if res else "")
                self.statusBar().showMessage("No data")
                return
            fields, rows = res
            for f in (L.F_TYPE, L.F_SUBZ):
                if f not in fields:
                    raise ValueError(f"Field '{f}' not found in the layer - set it under Field names.")
            lyr = self.layer_cb.currentLayer()
            cfield = self.color_field.currentField() or L.F_SUBZ
            self.colors = LC.build(rows, fields, cfield, self.cfg, self.lib,
                                   LC.layer_colors(lyr, cfield) if lyr else {}, L.BASE_COLORS, L.F_SUBZ, L.F_AREA)
            L.set_palette(self.colors["groups"])
            self.fill_colour_tab(rows)
            S = L.summarise(rows, self.effective_site_total() or None)
            self.data = (fields, rows, S)
            for pane, scene in zip(self.slide_panes, L.build_slides(S)):
                pane.view.set_scene(scene)
            self.fill_tables(rows, S)
            sel = " (selection)" if self.sel_only.isChecked() else ""
            self.count_lbl.setText(f"{len(rows)} polygons{sel}")
            self.statusBar().showMessage(
                f"{len(rows)} polygons · data {S['shp_total']:,.0f} sqm · grand total {S['total']:,.0f} sqm"
                f" · Road adjusted {S['road_adj']:+,.2f}")
        except Exception as e:  # keep the window alive on bad input
            self.statusBar().showMessage("Error: " + str(e))
            self.say("Error: " + str(e))

    def fill_tables(self, rows, S):
        T = S["total"] or 1.0
        fi, fd, fp = L.fmt_int, L.fmt_dec, L.fmt_pct

        def line(lab, v, **st):
            return ([lab, fi(v), fd(v / 1e4), fp(v / T)], st)

        mp = []
        for hdr, d, hx in (("SALEABLE", S["sale"], L.SALEABLE_HDR), ("NON-SALEABLE", S["nons"], L.NON_SALEABLE_HDR)):
            mp.append(line(hdr, sum(d.values()), rowfill=hx, bold=True))
            mp += [line(k, v, fill=L.color_of(k)) for k, v in d.items()]
            mp.append(([""] * 4, {}))
        mp.append(line("GRAND TOTAL", S["total"], rowfill="000000", bold=True))
        mp.append(([""] * 4, {}))
        mp.append((["Total in data", fi(S["shp_total"]), "", ""], {}))
        mp.append((["Adjustment added to Road", fd(S["road_adj"]), "", ""], {}))
        fill_table(self.tables["Overall MP"], ["Land use", "Area (sqm)", "Area (Ha)", "%"], mp)

        it = sum(S["ind"].values()) or 1.0
        ind = [([k, fi(v), fd(v / 1e4), fp(v / it), fp(v / T)], {"fill": L.color_of(k)}) for k, v in S["ind"].items()]
        ind.append((["Industrial Total", fi(sum(S["ind"].values())), fd(sum(S["ind"].values()) / 1e4),
                     fp(1 if S["ind"] else 0), fp(sum(S["ind"].values()) / T)], {"bold": True}))
        fill_table(self.tables["Industrial"], ["Industrial type", "Area (sqm)", "Area (Ha)", "% of Industrial", "% of Overall"], ind)

        pop = [([k, fi(v)], {}) for k, v in S["pop"].items()]
        pop.append((["Grand Total", fi(sum(S["pop"].values()))], {"bold": True}))
        fill_table(self.tables["Population"], ["Population type", "Population"], pop)

        ph = []
        for p, subs in L.phasing_data(rows, S).items():
            tot = sum(subs.values())
            ph.append(([p.upper(), "", fi(tot), fd(tot / 1e4), fp(tot / T)], {"rowfill": L.SUBTOTAL_GREY, "bold": True}))
            ph += [(["", k, fi(v), fd(v / 1e4), fp(v / T)], {}) for k, v in subs.items()]
            g, r = subs.get("Green", 0.0), subs.get(L.ROAD, 0.0)
            ph.append((["", "→ Green / Parcel / Road", "", "", f"{fp(g / T)} / {fp((tot - g - r) / T)} / {fp(r / T)}"], {}))
        fill_table(self.tables["Phasing"], ["Phase", "SUBZONA", "Area (sqm)", "Area (Ha)", "% of site"], ph)
        for i in range(self.tables["Phasing"].rowCount()):
            item = self.tables["Phasing"].item(i, 1)
            if item and item.text() in L.COLORS:
                hx = L.color_of(item.text())
                item.setBackground(QColor("#" + hx))
                item.setForeground(QColor("#FFFFFF" if L.is_dark(hx) else "#000000"))

        def n(v, d=None):
            if v in (None, ""):
                return ""
            return fi(v) if d is None else fd(v, d)

        ins = []
        for r in L.intensity_rows(rows):
            ins.append(([r["Type"], r["SUBZONA"], r["Parcel_ID"], r["Phase"] or "", r["BCR"] or "", r["FAR"] or "",
                         r["MaxH"] or "", r["GOSR"] or "", n(r["Area"]), n(r["Area"] / 1e4, 4), n(r["GFA"]),
                         n(r["Footprint"]), n(r["Pop"])], {}))
        fill_table(self.tables["Intensitas"], ["Type", "SUBZONA", "Parcel_ID", "Phase", "BCR (%)", "FAR", "Max Height",
                                               "GOSR (%)", "Area (sqm)", "Area (Ha)", "GFA (sqm)", "Max Footprint", "Pop"], ins)
        tw = self.tables["Intensitas"]
        for i in range(tw.rowCount()):
            item = tw.item(i, 1)
            if item:
                hx = L.color_of(item.text())
                item.setBackground(QColor("#" + hx))
                item.setForeground(QColor("#FFFFFF" if L.is_dark(hx) else "#000000"))

        checks = L.data_checks(rows, S)
        fill_table(self.tables["Checks"], ["Check", "Result"],
                   [([k, v], {"fill": "FFE699"} if (i >= 4 and v != "OK") else {}) for i, (k, v) in enumerate(checks)])
        warn = sum(1 for k, v in checks[4:] if v != "OK")
        self.tabs.setTabText(self.tabs.indexOf(self.tables["Checks"]), f"Checks ⚠ {warn}" if warn else "Checks ✓")

    # ---- output -------------------------------------------------------------
    def pick_file(self, edit, filt):
        p, _ = QFileDialog.getOpenFileName(self, "Choose file", edit.text(), filt)
        if p:
            edit.setText(p)

    def pick_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Output folder", self.out_dir.text())
        if d:
            self.out_dir.setText(d)

    def open_path(self, p):
        if p and os.path.exists(p):
            QDesktopServices.openUrl(QUrl.fromLocalFile(p))
        else:
            self.statusBar().showMessage("Nothing to open yet - generate the files first")

    def generate(self):
        self.refresh()
        if not self.data:
            QMessageBox.warning(self, APP_NAME, "No data to export - choose a land use layer.")
            return
        fields, rows, S = self.data
        lyr = self.layer_cb.currentLayer()
        base = self.base_name.text().strip() or (lyr.name() if lyr else "LUP")
        if self.sel_only.isChecked() and not self.base_name.text().strip():
            base += "_selection"
        out = self.out_dir.text().strip() or os.path.join(HERE, "Output")
        xlsx = os.path.join(out, f"{base}_LUP_Database.xlsx") if self.want_xlsx.isChecked() else None
        pptx = os.path.join(out, f"{base}_LUP_Tables.pptx") if self.want_pptx.isChecked() else None
        if not (xlsx or pptx):
            QMessageBox.warning(self, APP_NAME, "Tick at least one output (Excel or PowerPoint).")
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            self.say("-" * 60)
            L.run(fields, rows, xlsx, pptx, self.effective_site_total(), self.template.text() or None,
                  ["id", "en"][self.style.currentIndex()], log=self.say, colors=self.colors,
                  pivots=self.want_pivots.isChecked())
            self.last_outputs = (xlsx, pptx)
            self.statusBar().showMessage("Saved to " + out)
            self.say("Done.")
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, APP_NAME, str(e))
            self.say("Error: " + str(e))
            return
        QApplication.restoreOverrideCursor()
        if self.open_after.isChecked():
            for p in (xlsx, pptx):
                self.open_path(p)


def main():
    """Standalone entry point (run with QGIS's Python, see Run_LUP_Widget.bat)."""
    from qgis.core import QgsApplication
    prefix = os.environ.get("QGIS_PREFIX_PATH", "")
    app = QgsApplication([], True)
    if prefix:
        QgsApplication.setPrefixPath(prefix, True)
    app.initQgis()
    w = LupWindow()
    w.show()
    code = app.exec_()
    app.exitQgis()
    return code


if __name__ == "__main__":
    sys.exit(main())
