"""Project Setup window: drives qgis_project_setup_r2.py (imported, never run as
a script). The right-hand tree is both the preview of the layer tree and the
choice of datasets to add."""

import os

from qgis.PyQt.QtCore import Qt
from qgis.PyQt.QtGui import QBrush, QColor
from qgis.PyQt.QtWidgets import (QButtonGroup, QCheckBox, QFileDialog, QFormLayout, QGroupBox,
                                 QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton,
                                 QRadioButton, QToolButton, QTreeWidget, QTreeWidgetItem,
                                 QVBoxLayout)
from qgis.core import QgsProject

from .tool_window import ToolWindow, load_script

MODULE = "qgis_project_setup"
BASEMAP_GROUP = "Basemaps (loose layers at the bottom)"
REPORT_TO = "the log in the Project Setup window"
GREEN, GREY, AMBER = "#2E7D32", "#888888", "#B26A00"


class SetupWindow(ToolWindow):
    SETTINGS_KEY = "project_setup"

    def __init__(self, iface, script_path, parent=None):
        super().__init__(iface, parent, "Project Setup - layer tree, basemaps and live datasets")
        self.script_path = script_path
        self.mod = load_script(script_path, MODULE)
        self.items = {}      # dataset name -> tree item
        self.build_frame("Set up project")
        self.build_ui()
        self.restore_frame()
        self.restore_settings()
        self.connect_project()
        self.safe_refresh()

    # ---- UI ----------------------------------------------------------------
    def build_ui(self):
        m = self.mod
        g = QGroupBox("1. Project boundary")
        lay = QVBoxLayout(g)
        self.b_keep = QRadioButton("Keep the boundary in the project")
        self.b_keep.setToolTip(f"Use the '{m.BOUNDARY_NAME}' layer already in the project")
        self.b_file = QRadioButton("Add a boundary shapefile:")
        self.b_none = QRadioButton("No boundary")
        self.b_group = QButtonGroup(self)
        for k, b in enumerate((self.b_keep, self.b_file, self.b_none)):
            self.b_group.addButton(b, k)
            lay.addWidget(b)
            b.toggled.connect(self.schedule)
        row = QHBoxLayout()
        self.b_path = QLineEdit()
        self.b_path.setPlaceholderText("project boundary .shp")
        self.b_path.textChanged.connect(lambda _: self.b_file.setChecked(True))
        tb = QToolButton()
        tb.setText("…")
        tb.clicked.connect(self.browse_boundary)
        row.addWidget(self.b_path, 1)
        row.addWidget(tb)
        lay.addLayout(row)
        self.b_info = QLabel("")
        self.b_info.setWordWrap(True)
        self.b_info.setStyleSheet("color:#666")
        lay.addWidget(self.b_info)
        self.side.addWidget(g)

        g = QGroupBox("2. Datasets")
        f = QFormLayout(g)
        hint = QLabel("Tick the datasets to add in the <b>Layer tree</b> tab. Layers already in "
                      "the project are kept as they are.")
        hint.setWordWrap(True)
        f.addRow(hint)
        row = QHBoxLayout()
        for text, on in (("Tick all", True), ("Untick all", False)):
            b = QPushButton(text)
            b.clicked.connect(lambda _, on=on: self.set_all(on))
            row.addWidget(b)
        f.addRow(row)
        self.restyle = QCheckBox("Restyle layers added before")
        self.restyle.setToolTip("Apply the setup QML styles again to layers already in the project.\n"
                                "Off: manual style edits are safe.")
        f.addRow(self.restyle)
        self.side.addWidget(g)

        g = QGroupBox("3. What happens")
        lay = QVBoxLayout(g)
        info = QLabel("Builds the missing groups, adds the boundary (and its bookmark), the ticked "
                      "basemaps and live datasets (switched off, nothing downloaded) and the OSM "
                      "local roads, applies the setup styles, then sets the project CRS to the "
                      "boundary and zooms to it at a golden scale. Safe to run again.")
        info.setWordWrap(True)
        info.setStyleSheet("color:#444")
        lay.addWidget(info)
        self.side.addWidget(g)
        self.finish_side()

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Layer tree", "Status", "Source"])
        self.tree.setColumnWidth(0, 300)
        self.tree.setColumnWidth(1, 190)
        self.tree.itemChanged.connect(self.on_item_changed)
        self.tabs.addTab(self.tree, "Layer tree")
        self.build_tree()

    def build_tree(self):
        m = self.mod
        self.tree.blockSignals(True)
        self.tree.clear()
        self.items = {}
        self.group_items = {}
        datasets = {}
        for group, name, url, _ in m.VECTOR_LAYERS:
            datasets.setdefault(group, []).append((name, url))
        for name, slug in m.RISK_LAYERS:
            datasets.setdefault(m.RISK_GROUP, []).append((name, f"InaRISK {slug}"))
        datasets.setdefault("Connectivity", []).append((m.LOCAL_ROADS_NAME, "OpenStreetMap (Overpass)"))
        for name in m.ROOT_GROUPS:
            if name == m.SEPARATOR:
                sep = QTreeWidgetItem(self.tree, ["────────────", "", ""])
                sep.setFlags(Qt.NoItemFlags)
                continue
            gi = QTreeWidgetItem(self.tree, [name, "", ""])
            f = gi.font(0)
            f.setBold(True)
            gi.setFont(0, f)
            self.group_items[name] = gi
            if name == m.BOUNDARY_GROUP:
                bi = QTreeWidgetItem(gi, [m.BOUNDARY_NAME, "", "shapefile (step 1)"])
                self.items[m.BOUNDARY_NAME] = bi
            entries = datasets.get(name, [])
            if entries:
                gi.setFlags(gi.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsAutoTristate)
            for ds, src in entries:
                it = QTreeWidgetItem(gi, [ds, "", src])
                it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
                it.setCheckState(0, Qt.Checked)
                self.items[ds] = it
            gi.setExpanded(bool(entries))
        gi = QTreeWidgetItem(self.tree, [BASEMAP_GROUP, "", ""])
        f = gi.font(0)
        f.setBold(True)
        gi.setFont(0, f)
        gi.setFlags(gi.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsAutoTristate)
        for name, url, _ in m.BASEMAPS:
            it = QTreeWidgetItem(gi, [name + ("  (switched on)" if name == m.DEFAULT_BASEMAP else ""), "", url])
            it.setData(0, Qt.UserRole, name)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(0, Qt.Checked)
            self.items[name] = it
        gi.setExpanded(True)
        self.tree.blockSignals(False)

    # ---- settings ----------------------------------------------------------------
    def restore_settings(self):
        st = self.settings
        self.restyle.setChecked(st.value("restyle", False, type=bool))
        self.b_path.blockSignals(True)
        self.b_path.setText(st.value("boundary_path", "") or "")
        self.b_path.blockSignals(False)
        unticked = set(st.value("unticked", []) or [])
        self.tree.blockSignals(True)
        for name, it in self.items.items():
            if it.flags() & Qt.ItemIsUserCheckable:
                it.setCheckState(0, Qt.Unchecked if name in unticked else Qt.Checked)
        self.tree.blockSignals(False)
        has = self.has_boundary()
        (self.b_keep if has else (self.b_file if self.b_path.text() else self.b_none)).setChecked(True)

    def save_settings(self):
        st = self.settings
        st.setValue("restyle", self.restyle.isChecked())
        st.setValue("boundary_path", self.b_path.text())
        st.setValue("unticked", [n for n, it in self.items.items()
                                 if it.flags() & Qt.ItemIsUserCheckable and it.checkState(0) != Qt.Checked])

    # ---- project watching ------------------------------------------------------------
    def connect_project(self):
        p = QgsProject.instance()
        self._sigs = [(p.layersAdded, self.schedule), (p.layersRemoved, self.schedule),
                      (p.layerTreeRoot().addedChildren, self.schedule),
                      (p.layerTreeRoot().removedChildren, self.schedule),
                      (p.readProject, self.schedule), (p.projectSaved, self.schedule)]
        for sig, slot in self._sigs:
            sig.connect(slot)

    def disconnect_project(self):
        for sig, slot in getattr(self, "_sigs", []):
            try:
                sig.disconnect(slot)
            except (TypeError, RuntimeError):
                pass
        self._sigs = []

    # ---- helpers ----------------------------------------------------------------
    def browse_boundary(self):
        start = self.b_path.text() or QgsProject.instance().homePath()
        path, _ = QFileDialog.getOpenFileName(self, "Select the project boundary shapefile", start,
                                              "Shapefiles (*.shp)")
        if path:
            self.b_path.setText(path)
            self.b_file.setChecked(True)

    def set_all(self, on):
        self.tree.blockSignals(True)
        for it in self.items.values():
            if it.flags() & Qt.ItemIsUserCheckable:
                it.setCheckState(0, Qt.Checked if on else Qt.Unchecked)
        self.tree.blockSignals(False)
        self.safe_refresh()

    def on_item_changed(self, item, col):
        if col == 0:
            self.schedule()

    def ticked(self, name):
        it = self.items.get(name)
        return it is not None and it.checkState(0) == Qt.Checked

    def has_boundary(self):
        """True when the setup's boundary layer sits in its group (the script keeps it)."""
        return self.exists(self.mod.BOUNDARY_GROUP, self.mod.BOUNDARY_NAME)

    def exists(self, group, name):
        m = self.mod
        root = QgsProject.instance().layerTreeRoot()
        if group is None:
            return m.find_layer_index(root, name) is not None
        groups = m.find_groups(root, group)
        return bool(groups) and m.find_layer_index(groups[0], name) is not None

    @staticmethod
    def set_status(it, text, color):
        it.setText(1, text)
        it.setForeground(1, QBrush(QColor(color)))

    # ---- preview --------------------------------------------------------------------
    def refresh(self):
        m = self.mod
        project = QgsProject.instance()
        has_boundary = self.has_boundary()
        self.b_keep.setEnabled(has_boundary)
        self.b_none.setEnabled(not has_boundary)  # an existing boundary is always kept
        if has_boundary and self.b_none.isChecked():
            self.b_keep.setChecked(True)
        if has_boundary:
            self.b_info.setText(f"'{m.BOUNDARY_NAME}' is in the project. To use another boundary, "
                                "choose a shapefile: the current one is replaced.")
        else:
            self.b_info.setText(f"No '{m.BOUNDARY_NAME}' yet. Without one, local roads use the map view "
                                "and the map is not zoomed to a site.")
            if self.b_keep.isChecked():
                self.b_none.setChecked(True)

        self.tree.blockSignals(True)
        add = keep = 0
        root = project.layerTreeRoot()
        for name, gi in self.group_items.items():
            there = bool(m.find_groups(root, name))
            self.set_status(gi, "group in project" if there else "group will be created",
                            GREY if there else GREEN)
        bi = self.items[m.BOUNDARY_NAME]
        if self.b_file.isChecked() and self.b_path.text():
            ok = os.path.isfile(self.b_path.text())
            self.set_status(bi, ("will be replaced" if has_boundary else "will be added") if ok
                            else "shapefile not found", GREEN if ok else AMBER)
            add += ok
        elif has_boundary:
            self.set_status(bi, "in project (kept)", GREY)
            keep += 1
        else:
            self.set_status(bi, "none", GREY)
        for group, name, _, _ in m.VECTOR_LAYERS:
            a, k = self.status_for(group, name)
            add, keep = add + a, keep + k
        for name, _ in m.RISK_LAYERS:
            a, k = self.status_for(m.RISK_GROUP, name)
            add, keep = add + a, keep + k
        for name, _, _ in m.BASEMAPS:
            a, k = self.status_for(None, name)
            add, keep = add + a, keep + k
        it = self.items[m.LOCAL_ROADS_NAME]
        if self.exists("Connectivity", m.LOCAL_ROADS_NAME):
            self.set_status(it, "in project (kept)", GREY)
            keep += 1
        elif not self.ticked(m.LOCAL_ROADS_NAME):
            self.set_status(it, "not ticked", GREY)
        elif not project.fileName():
            self.set_status(it, "save the project first", AMBER)
        else:
            self.set_status(it, "will be downloaded", GREEN)
            add += 1
        self.tree.blockSignals(False)
        self.statusBar().showMessage(f"{add} layer(s) to add, {keep} already in the project")

    def status_for(self, group, name):
        it = self.items[name]
        if self.exists(group, name):
            self.set_status(it, "in project (kept)", GREY)
            return 0, 1
        if not self.ticked(name):
            self.set_status(it, "not ticked", GREY)
            return 0, 0
        self.set_status(it, "will be added", GREEN)
        return 1, 0

    # ---- run ----------------------------------------------------------------------
    def run(self):
        if self.running:
            return
        m = load_script(self.script_path, MODULE)  # fresh copy: config edits below stay local
        project = QgsProject.instance()
        boundary_path = ""
        if self.b_file.isChecked():
            boundary_path = self.b_path.text().strip()
            if not os.path.isfile(boundary_path):
                QMessageBox.warning(self, "Project Setup", "Choose an existing boundary shapefile, "
                                    "or pick another boundary option.")
                return
            old = project.mapLayersByName(m.BOUNDARY_NAME)
            if old:
                if QMessageBox.question(self, "Project Setup",
                                        f"Replace the '{m.BOUNDARY_NAME}' layer in the project with\n"
                                        f"{boundary_path}?") != QMessageBox.Yes:
                    return
                project.removeMapLayers([l.id() for l in old])
        m.RESTYLE_EXISTING = self.restyle.isChecked()
        m.VECTOR_LAYERS = [v for v in m.VECTOR_LAYERS if self.ticked(v[1])]
        m.RISK_LAYERS = [v for v in m.RISK_LAYERS if self.ticked(v[0])]
        m.BASEMAPS = [v for v in m.BASEMAPS if self.ticked(v[0])]
        self.save_settings()
        self.run_logged("Project Setup", m.main, boundary_path=boundary_path,
                        local_roads=self.ticked(m.LOCAL_ROADS_NAME), report_to=REPORT_TO)
        self.safe_refresh()
