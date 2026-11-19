"""Export Pivot to Excel window: live, colour-coded preview of the pivot table for a
vector layer (following selection, edits and style changes), then export to .xlsx.
Same layout as the Export Land Use window. The data comes from vpe_layer.collect()
and the table from vpe_core.pivot_grid(), the very code the Excel export uses."""

import json
import os
import traceback

from qgis.PyQt.QtCore import QSettings, Qt, QTimer
from qgis.PyQt.QtGui import QColor, QKeySequence
from qgis.PyQt.QtWidgets import (QAbstractItemView, QApplication, QCheckBox, QComboBox, QFileDialog,
                                 QFormLayout, QFrame, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
                                 QListWidget, QListWidgetItem, QMainWindow, QPlainTextEdit, QPushButton,
                                 QScrollArea, QShortcut, QSizePolicy, QSplitter, QTableWidget,
                                 QTableWidgetItem, QTabWidget, QToolButton, QVBoxLayout, QWidget)
from qgis.core import QgsMapLayerProxyModel, QgsProject, QgsVectorLayer
from qgis.gui import QgsMapLayerComboBox

from . import vpe_core as core
from . import vpe_layer as VL

APP_NAME = "Export Pivot to Excel"
SYM_LABEL = "‹Symbology class›"
SOURCE_PREVIEW_ROWS = 1000


def make_table():
    t = QTableWidget()
    t.setEditTriggers(QTableWidget.NoEditTriggers)
    t.verticalHeader().setVisible(False)
    t.setWordWrap(False)
    t.verticalHeader().setDefaultSectionSize(t.fontMetrics().height() + 8)
    return t


def cell(text, fill=None, bold=False, right=False, header=False):
    it = QTableWidgetItem("" if text is None else str(text))
    if header and not fill:
        fill = core.HDR
    if fill:
        it.setBackground(QColor("#" + fill))
        it.setForeground(QColor("#FFFFFF" if core.is_dark(fill) else "#000000"))
    if bold or header:
        f = it.font()
        f.setBold(True)
        it.setFont(f)
    if header:
        it.setTextAlignment(Qt.AlignCenter)
    elif right:
        it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
    return it


def fit_columns(t, lo=60, hi=320):
    t.resizeColumnsToContents()
    for j in range(t.columnCount()):
        t.setColumnWidth(j, max(lo, min(hi, t.columnWidth(j) + 14)))


class PivotWindow(QMainWindow):
    def __init__(self, iface=None, parent=None):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() | Qt.Window)
        self.iface = iface
        self.settings = QSettings("UrbanDesignSetup", "vector_pivot")
        self.layer = None
        self.last_output = None
        self.setWindowTitle(APP_NAME + " - pivot table coloured by the layer symbology")
        self.setMinimumSize(640, 420)
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.setInterval(350)
        self.timer.timeout.connect(self.refresh)
        self.build_ui()
        self.restore()
        QShortcut(QKeySequence("F5"), self, activated=self.refresh)
        QShortcut(QKeySequence("Ctrl+S"), self, activated=self.export)

    # ---- UI ---------------------------------------------------------------
    def build_ui(self):
        host = QWidget()
        side = QVBoxLayout(host)

        g = QGroupBox("1. Layer")
        f = QFormLayout(g)
        self.layer_cb = QgsMapLayerComboBox()
        self.layer_cb.setFilters(QgsMapLayerProxyModel.VectorLayer)
        self.layer_cb.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.layer_cb.setMinimumContentsLength(8)
        self.layer_cb.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.layer_cb.layerChanged.connect(self.on_layer_changed)
        browse = QPushButton("Open file…")
        browse.clicked.connect(self.browse_layer)
        row = QHBoxLayout()
        row.addWidget(self.layer_cb, 1)
        row.addWidget(browse)
        f.addRow(row)
        self.sel_only = QCheckBox("Selected features only")
        self.sel_only.toggled.connect(self.schedule)
        f.addRow(self.sel_only)
        self.auto = QCheckBox("Live preview")
        self.auto.setToolTip("Refresh the preview when the layer's features, selection or style change")
        self.auto.setChecked(True)
        f.addRow(self.auto)
        self.count_lbl = QLabel("")
        self.count_lbl.setStyleSheet("color:#666")
        f.addRow(self.count_lbl)
        side.addWidget(g)

        g = QGroupBox("2. Pivot")
        f = QFormLayout(g)
        lbl = QLabel("Rows - tick, drag to order (none ticked = the symbology field)")
        lbl.setWordWrap(True)
        f.addRow(lbl)
        self.rows_list = QListWidget()
        self.rows_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.rows_list.setMaximumHeight(150)
        self.rows_list.itemChanged.connect(self.schedule)
        self.rows_list.model().rowsMoved.connect(self.schedule)
        f.addRow(self.rows_list)
        self.col_cb = QComboBox()
        self.col_cb.currentIndexChanged.connect(self.schedule)
        f.addRow("Column", self.col_cb)
        f.addRow(QLabel("Values (numeric fields)"))
        self.vals_list = QListWidget()
        self.vals_list.setMaximumHeight(110)
        self.vals_list.itemChanged.connect(self.schedule)
        f.addRow(self.vals_list)
        self.agg_cb = QComboBox()
        self.agg_cb.addItems(VL.AGG_OPTIONS)
        self.agg_cb.currentIndexChanged.connect(self.schedule)
        f.addRow("Aggregate", self.agg_cb)
        self.measure_cb = QComboBox()
        self.measure_cb.addItems([m[0] for m in VL.MEASURES])
        self.measure_cb.currentIndexChanged.connect(self.schedule)
        f.addRow("Geometry", self.measure_cb)
        self.count_chk = QCheckBox("Feature count")
        self.count_chk.setChecked(True)
        self.count_chk.toggled.connect(self.schedule)
        self.pct_chk = QCheckBox("% of total (sums and counts)")
        self.pct_chk.setChecked(True)
        self.pct_chk.toggled.connect(self.schedule)
        f.addRow(self.count_chk)
        f.addRow(self.pct_chk)
        side.addWidget(g)

        g = QGroupBox("3. Output")
        f = QFormLayout(g)
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
        self.pivot_cb = QComboBox()
        self.pivot_cb.addItems(["Excel pivot table if Excel is installed", "Static table (no Excel needed)",
                                "Excel pivot table only"])
        self.pivot_cb.setToolTip("\n".join(VL.PIVOT_OPTIONS))
        f.addRow("Pivot sheet", self.pivot_cb)
        self.open_after = QCheckBox("Open the workbook when done")
        self.open_after.setChecked(True)
        f.addRow(self.open_after)
        side.addWidget(g)

        self.export_btn = QPushButton("Export to Excel  (Ctrl+S)")
        self.export_btn.setMinimumHeight(34)
        self.export_btn.setStyleSheet("font-weight:bold")
        self.export_btn.clicked.connect(self.export)
        side.addWidget(self.export_btn)
        row = QHBoxLayout()
        ox = QPushButton("Open Excel")
        ox.clicked.connect(lambda: self.open_path(self.last_output))
        od = QPushButton("Open folder")
        od.clicked.connect(lambda: self.open_path(self.out_dir.text()))
        row.addWidget(ox)
        row.addWidget(od)
        side.addLayout(row)
        ref = QPushButton("Refresh preview  (F5)")
        ref.clicked.connect(self.refresh)
        side.addWidget(ref)
        side.addStretch(1)
        for cb in host.findChildren(QComboBox):  # combos shrink: the panel never outgrows its column
            cb.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            cb.setMinimumContentsLength(8)
            cb.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        left = QScrollArea()
        left.setWidgetResizable(True)
        left.setWidget(host)
        left.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        left.setMinimumWidth(280)
        left.setFrameShape(QFrame.NoFrame)

        self.tabs = QTabWidget()
        self.pivot_tbl = make_table()
        self.pivot_tbl.horizontalHeader().setVisible(False)
        self.colours_tbl = make_table()
        self.source_tbl = make_table()
        self.tabs.addTab(self.pivot_tbl, "Pivot")
        self.tabs.addTab(self.colours_tbl, "Colours")
        self.tabs.addTab(self.source_tbl, "Source")
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(2000)

        self.vsplit = QSplitter(Qt.Vertical)
        self.vsplit.addWidget(self.tabs)
        self.vsplit.addWidget(self.log)
        self.vsplit.setStretchFactor(0, 5)
        self.vsplit.setStretchFactor(1, 1)
        self.vsplit.setSizes([600, 110])
        self.hsplit = QSplitter(Qt.Horizontal)
        self.hsplit.addWidget(left)
        self.hsplit.addWidget(self.vsplit)
        self.hsplit.setStretchFactor(1, 1)
        self.hsplit.setSizes([420, 860])
        self.setCentralWidget(self.hsplit)
        self.statusBar().showMessage("Ready")

    # ---- settings -----------------------------------------------------------
    def restore(self):
        st = self.settings
        g = st.value("geometry")
        if g is not None:
            self.restoreGeometry(g)
        else:
            self.resize(1240, 800)
        for sp, key in ((self.hsplit, "hsplit"), (self.vsplit, "vsplit")):
            v = st.value(key)
            if v is not None:
                sp.restoreState(v)
        self.agg_cb.setCurrentIndex(int(st.value("agg", 0)))
        self.measure_cb.setCurrentIndex(int(st.value("measure", 0)))
        self.pivot_cb.setCurrentIndex(int(st.value("pivot", 0)))
        self.count_chk.setChecked(str(st.value("count", True)).lower() in ("true", "1"))
        self.pct_chk.setChecked(str(st.value("percent", True)).lower() in ("true", "1"))
        self.out_dir.setText(st.value("out_dir", "") or QgsProject.instance().homePath()
                             or os.path.expanduser("~"))
        self.on_layer_changed(self.layer_cb.currentLayer())

    def save_settings(self):
        st = self.settings
        st.setValue("geometry", self.saveGeometry())
        st.setValue("hsplit", self.hsplit.saveState())
        st.setValue("vsplit", self.vsplit.saveState())
        st.setValue("agg", self.agg_cb.currentIndex())
        st.setValue("measure", self.measure_cb.currentIndex())
        st.setValue("pivot", self.pivot_cb.currentIndex())
        st.setValue("count", self.count_chk.isChecked())
        st.setValue("percent", self.pct_chk.isChecked())
        st.setValue("out_dir", self.out_dir.text())
        self.save_layer_choice()

    def closeEvent(self, e):
        self.save_settings()
        self.watch(None)
        super().closeEvent(e)

    def layer_key(self):
        return "layers/" + (self.layer.name() if self.layer is not None else "")

    def save_layer_choice(self):
        if self.layer is None:
            return
        choice = dict(rows=self.chosen_rows(), col=self.col_cb.currentData(), vals=self.chosen_values(),
                      order=[self.rows_list.item(i).data(Qt.UserRole) for i in range(self.rows_list.count())])
        self.settings.setValue(self.layer_key(), json.dumps(choice))

    # ---- layer ---------------------------------------------------------------
    def set_layer(self, layer):
        self.layer_cb.setLayer(layer)

    def browse_layer(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open vector layer", self.settings.value("last_file", ""),
                                              "Vector files (*.shp *.gpkg *.geojson *.json *.kml *.tab);;All files (*)")
        if not path:
            return
        lyr = QgsVectorLayer(path, os.path.splitext(os.path.basename(path))[0], "ogr")
        if not lyr.isValid():
            self.say(f"Could not open {path}")
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
        self.save_layer_choice()
        self.watch(layer)
        for w in (self.rows_list, self.vals_list, self.col_cb):
            w.blockSignals(True)
        self.rows_list.clear()
        self.vals_list.clear()
        self.col_cb.clear()
        self.col_cb.addItem("(none)", None)
        if layer is not None:
            self.base_name.setPlaceholderText(layer.name())
            saved = {}
            try:
                saved = json.loads(self.settings.value(self.layer_key(), "") or "{}")
            except ValueError:
                pass
            names = [VL.SYMBOLOGY] + [f.name() for f in layer.fields()]
            order = [n for n in saved.get("order", []) if n in names]
            names = order + [n for n in names if n not in order]
            style_field = VL.symbology_field(layer)
            rows = set(saved.get("rows") or [style_field or VL.SYMBOLOGY])
            for n in names:
                it = QListWidgetItem(SYM_LABEL if n == VL.SYMBOLOGY else n)
                it.setData(Qt.UserRole, n)
                it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
                it.setCheckState(Qt.Checked if n in rows else Qt.Unchecked)
                if n == style_field or n == VL.SYMBOLOGY:
                    it.setToolTip("The layer's symbology")
                self.rows_list.addItem(it)
            self.col_cb.addItem(SYM_LABEL, VL.SYMBOLOGY)
            for fld in layer.fields():
                self.col_cb.addItem(fld.name(), fld.name())
            k = self.col_cb.findData(saved.get("col"))
            self.col_cb.setCurrentIndex(max(0, k))
            vals = set(saved.get("vals") or [])
            for fld in layer.fields():
                if fld.isNumeric():
                    it = QListWidgetItem(fld.name())
                    it.setData(Qt.UserRole, fld.name())
                    it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
                    it.setCheckState(Qt.Checked if fld.name() in vals else Qt.Unchecked)
                    self.vals_list.addItem(it)
            if not self.vals_list.count():
                self.vals_list.addItem(QListWidgetItem("(no numeric fields)"))
        for w in (self.rows_list, self.vals_list, self.col_cb):
            w.blockSignals(False)
        self.refresh()

    def chosen_rows(self):
        return [self.rows_list.item(i).data(Qt.UserRole) for i in range(self.rows_list.count())
                if self.rows_list.item(i).checkState() == Qt.Checked]

    def chosen_values(self):
        return [self.vals_list.item(i).data(Qt.UserRole) for i in range(self.vals_list.count())
                if self.vals_list.item(i).data(Qt.UserRole) and self.vals_list.item(i).checkState() == Qt.Checked]

    # ---- data -------------------------------------------------------------------
    def schedule(self, *args):
        self.timer.start()

    def say(self, msg):
        self.log.appendPlainText(msg)
        self.log.ensureCursorVisible()

    def collect(self):
        p = QgsProject.instance()
        return VL.collect(self.layer, selected=self.sel_only.isChecked(), row_fields=self.chosen_rows(),
                          col_field=self.col_cb.currentData(), value_fields=self.chosen_values(),
                          agg=VL.AGGS[self.agg_cb.currentIndex()], measure_index=self.measure_cb.currentIndex(),
                          add_count=self.count_chk.isChecked(), percent=self.pct_chk.isChecked(),
                          ellipsoid=p.ellipsoid(), transform_context=p.transformContext())

    def refresh(self):
        layer = self.layer
        if layer is None:
            self.count_lbl.setText("Choose a vector layer")
            for t in (self.pivot_tbl, self.colours_tbl, self.source_tbl):
                t.clear()
                t.setRowCount(0)
            return
        n = layer.selectedFeatureCount() if self.sel_only.isChecked() else layer.featureCount()
        self.count_lbl.setText(f"{n:,} feature(s)" + (" selected" if self.sel_only.isChecked() else "")
                               + f" · style: {layer.renderer().type() if layer.renderer() else 'none'}")
        try:
            data = self.collect()
        except VL.CollectError as e:
            self.statusBar().showMessage(str(e))
            self.pivot_tbl.setRowCount(0)
            return
        except Exception:
            self.say("Preview failed:\n" + traceback.format_exc())
            return
        if not data["rows"]:
            self.statusBar().showMessage("No features to show")
            self.pivot_tbl.setRowCount(0)
            return
        G = core.pivot_grid(data["rows"], data["spec"], data["colors"], data["orders"])
        self.fill_pivot(G)
        self.fill_colours(data)
        self.fill_source(data)
        msg = f"{len(data['rows']):,} features · rows: {', '.join(data['spec']['row_fields'])}"
        if data["spec"]["col_field"]:
            msg += f" · column: {data['spec']['col_field']}"
        for w in data["warnings"]:
            msg += " · " + w
        self.statusBar().showMessage(msg)

    def fill_pivot(self, G):
        t = self.pivot_tbl
        rf, cf = G["row_fields"], G["col_field"]
        nl = len(rf)
        nvals = len(G["lines"][0][2]) if G["lines"] else 0
        head = 2 if cf else 1
        t.clear()
        t.clearSpans()
        t.setColumnCount(nl + nvals)
        t.setRowCount(head + len(G["lines"]))
        for j, f in enumerate(rf):
            t.setItem(0, j, cell(f, header=True))
            if cf:
                t.setItem(1, j, cell("", header=True))
                t.setSpan(0, j, 2, 1)
        c = nl
        for cap, span in G["header_top"]:
            span = span if cf else 1
            t.setItem(0, c, cell(cap, header=True))
            if span > 1:
                t.setSpan(0, c, 1, span)
            c += span
        for k, (text, hx) in enumerate(G["header_sub"]):
            t.setItem(1, nl + k, cell(text, fill=hx, header=True))
        for i, (kind, labels, vals) in enumerate(G["lines"]):
            r = head + i
            bold = kind != "row"
            fill = {"subtotal": core.SUB, "total": "D9D9D9"}.get(kind)
            for j, (text, hx) in enumerate(labels):
                if kind == "total":
                    t.setItem(r, j, cell(text, header=True))
                else:
                    t.setItem(r, j, cell(text, fill=hx or (fill if j > 0 else None), bold=bold))
            for k, (v, numfmt) in enumerate(vals):
                t.setItem(r, nl + k, cell(core.format_value(v, numfmt), fill=fill, bold=bold, right=True))
        fit_columns(t, hi=360)

    def fill_colours(self, data):
        t = self.colours_tbl
        rows = []
        for f, scheme in data["colors"].items():
            counts = {}
            for r in data["rows"]:
                k = core.key(r.get(f))
                counts[k] = counts.get(k, 0) + 1
            for v in core.sort_values(list(scheme), data["orders"].get(f)):
                rows.append((f, v, scheme[v], counts.get(v, 0)))
        t.clear()
        t.setColumnCount(4)
        t.setHorizontalHeaderLabels(["Field", "Value", "HEX", "Features"])
        t.setRowCount(len(rows))
        for i, (f, v, hx, n) in enumerate(rows):
            t.setItem(i, 0, cell(f))
            t.setItem(i, 1, cell(v, fill=hx))
            t.setItem(i, 2, cell("#" + hx))
            t.setItem(i, 3, cell(f"{n:,}", right=True))
        fit_columns(t)

    def fill_source(self, data):
        t = self.source_tbl
        fields, rows = data["fields"], data["rows"][:SOURCE_PREVIEW_ROWS]
        t.clear()
        t.setColumnCount(len(fields))
        t.setHorizontalHeaderLabels(fields)
        t.setRowCount(len(rows))
        for i, r in enumerate(rows):
            for j, f in enumerate(fields):
                v = r.get(f)
                hx = (data["colors"].get(f) or {}).get(core.key(v))
                num = isinstance(v, (int, float)) and not isinstance(v, bool)
                t.setItem(i, j, cell(core.format_value(v, "#,##0.00") if isinstance(v, float) else v,
                                     fill=hx, right=num))
        fit_columns(t, hi=260)
        if len(data["rows"]) > SOURCE_PREVIEW_ROWS:
            self.say(f"Source tab shows the first {SOURCE_PREVIEW_ROWS:,} of {len(data['rows']):,} rows "
                     "(the workbook gets all of them).")

    # ---- export ------------------------------------------------------------------
    def pick_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Output folder", self.out_dir.text())
        if d:
            self.out_dir.setText(d)

    def open_path(self, p):
        if not p or not os.path.exists(p):
            self.say("Nothing to open yet.")
            return
        from .vpe_algorithm import open_file
        open_file(p)

    def export(self):
        if self.layer is None:
            self.say("Choose a vector layer first.")
            return
        folder = self.out_dir.text().strip() or os.path.expanduser("~")
        name = (self.base_name.text().strip() or self.layer.name()).replace("/", "_").replace("\\", "_")
        if not name.lower().endswith(".xlsx"):
            name += ".xlsx"
        path = os.path.join(folder, name)
        os.makedirs(folder, exist_ok=True)
        self.export_btn.setEnabled(False)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            data = self.collect()
            if not data["rows"]:
                self.say("No features to export.")
                return
            self.say(f"Writing {len(data['rows']):,} features to {path}")
            made = core.write_workbook(path, data["fields"], data["rows"], data["spec"], data["colors"],
                                       data["orders"], native=VL.PIVOT_KEYS[self.pivot_cb.currentIndex()],
                                       title=data["title"], note=data["note"], log=self.say)
            self.last_output = path
            self.say("Done - pivot sheet: " + ("Excel pivot table" if made == "excel" else "static table"))
            self.statusBar().showMessage(f"Exported {os.path.basename(path)}", 8000)
            self.save_settings()
            if self.open_after.isChecked():
                self.open_path(path)
        except PermissionError:
            self.say(f"Cannot write {path} - is it open in Excel?")
        except Exception as e:
            self.say(f"Export failed: {type(e).__name__}: {e}")
        finally:
            QApplication.restoreOverrideCursor()
            self.export_btn.setEnabled(True)
