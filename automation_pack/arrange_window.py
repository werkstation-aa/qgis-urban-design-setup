"""Project Arrange window: drives qgis_project_arrange.py (imported, never run as
a script). The preview tabs show the moves and styles the script plans, from its
own plan() and styles_for(), before anything changes."""

from qgis.PyQt.QtWidgets import QCheckBox, QFormLayout, QGroupBox, QLabel, QVBoxLayout
from qgis.core import QgsLayerTree, QgsMapLayerProxyModel, QgsProject, QgsWkbTypes
from qgis.gui import QgsMapLayerComboBox

from .tool_window import ToolWindow, fill_table, load_script, make_table, table_item

MODULE = "qgis_project_arrange"
REPORT_TO = "the log in the Project Arrange window"
OPTIONS = [  # (module constant, label, tooltip)
    ("POINTS_TO_NODES", "Point layers to Nodes",
     "Every point layer moves to Nodes, at the top of the tree, so it draws above everything else\n"
     "(except those in Top Overlays, Contour and Topography)."),
    ("SORT_GROUPS", "Sort each group",
     "Points, lines, polygons, rasters; within each geometry, layers follow the order\n"
     "of the types of information."),
    ("APPLY_STYLES", "Apply the setup styles",
     "Styles matching layers with the setup script's QML styles.\n"
     "The previous style stays available as 'Original' in the layer's Styles menu."),
    ("RESTYLE_EXISTING", "Restyle layers styled before",
     "Apply the styles again to layers an earlier run already styled."),
]


class ArrangeWindow(ToolWindow):
    SETTINGS_KEY = "project_arrange"

    def __init__(self, iface, script_path, styles_dir, parent=None):
        super().__init__(iface, parent, "Project Arrange - sort layers into the project setup tree")
        self.script_path = script_path
        self.styles_dir = styles_dir
        self.mod = self.load()
        self.build_frame("Arrange project")
        self.build_ui()
        self.restore_frame()
        self.restore_settings()
        self.connect_project()
        self.safe_refresh()

    def load(self):
        m = load_script(self.script_path, MODULE)
        m.STYLES_DIR = self.styles_dir  # the QML styles shipped with the plugin
        return m

    # ---- UI ----------------------------------------------------------------
    def build_ui(self):
        m = self.mod
        g = QGroupBox("1. Project boundary")
        f = QFormLayout(g)
        self.boundary = QgsMapLayerComboBox()
        self.boundary.setFilters(QgsMapLayerProxyModel.VectorLayer)
        self.boundary.setAllowEmptyLayer(True)
        try:
            self.boundary.setAllowEmptyLayer(True, "No boundary")
        except TypeError:
            pass
        self.boundary.layerChanged.connect(self.schedule)
        f.addRow(self.boundary)
        info = QLabel(f"Goes first in {m.BOUNDARY_GROUP}, is saved as the '{m.BOUNDARY_NAME}' bookmark, "
                      "and the map zooms to it at a golden scale.")
        info.setWordWrap(True)
        info.setStyleSheet("color:#666")
        f.addRow(info)
        self.side.addWidget(g)

        g = QGroupBox("2. Options")
        lay = QVBoxLayout(g)
        self.opts = {}
        for const, label, tip in OPTIONS:
            cb = QCheckBox(label)
            cb.setToolTip(tip)
            cb.setChecked(bool(getattr(m, const)))
            cb.toggled.connect(self.schedule)
            self.opts[const] = cb
            lay.addWidget(cb)
        self.side.addWidget(g)

        g = QGroupBox("3. What happens")
        lay = QVBoxLayout(g)
        info = QLabel("Every layer and group is moved to the setup group its name points to "
                      "(see the Planned moves tab), the setup tree is rebuilt around them with "
                      "dividers and basemaps last, and empty groups are removed. Nothing is removed "
                      "from the project; visible layers stay visible.")
        info.setWordWrap(True)
        info.setStyleSheet("color:#444")
        lay.addWidget(info)
        self.side.addWidget(g)
        self.finish_side()

        self.moves_tbl = make_table()
        self.styles_tbl = make_table()
        self.tabs.addTab(self.moves_tbl, "Planned moves")
        self.tabs.addTab(self.styles_tbl, "Styles")

    # ---- settings ----------------------------------------------------------------
    def restore_settings(self):
        for const, cb in self.opts.items():
            v = self.settings.value(const)
            if v is not None:
                cb.setChecked(str(v).lower() in ("true", "1"))
        project = QgsProject.instance()
        vectors, pre = self.mod.boundary_candidates(project)
        if vectors:
            self.boundary.setLayer(vectors[pre])
        else:
            self.boundary.setLayer(None)

    def save_settings(self):
        for const, cb in self.opts.items():
            self.settings.setValue(const, cb.isChecked())

    # ---- project watching ------------------------------------------------------------
    def connect_project(self):
        p = QgsProject.instance()
        root = p.layerTreeRoot()
        self._sigs = [(p.layersAdded, self.schedule), (p.layersRemoved, self.schedule),
                      (root.addedChildren, self.schedule), (root.removedChildren, self.schedule),
                      (root.nameChanged, self.schedule), (p.readProject, self.on_project_read)]
        for sig, slot in self._sigs:
            sig.connect(slot)

    def disconnect_project(self):
        for sig, slot in getattr(self, "_sigs", []):
            try:
                sig.disconnect(slot)
            except (TypeError, RuntimeError):
                pass
        self._sigs = []

    def on_project_read(self, *args):
        self.restore_settings()
        self.schedule()

    # ---- preview --------------------------------------------------------------------
    def configure(self, m):
        for const, cb in self.opts.items():
            setattr(m, const, cb.isChecked())
        m.DRY_RUN = False
        layer = self.boundary.currentLayer()
        m.boundary_id = layer.id() if layer is not None else None
        return layer

    def refresh(self):
        m = self.mod
        self.configure(m)
        project = QgsProject.instance()
        moves, emptied = m.plan(project.layerTreeRoot())
        rows = []
        for node, dest in moves:
            group = QgsLayerTree.isGroup(node)
            parent = node.parent()
            src = parent.name() if parent is not None and parent.parent() is not None else "(top level)"
            theme = m.theme_of(node.name())
            dest_txt = {m.BASEMAP: "top level (web basemap, bottom)",
                        m.OTHER_RASTER: "top level (raster, above Orthophoto)"}.get(dest, dest)
            kind = "group"
            if not group and node.layer() is not None:
                lyr = node.layer()
                kind = (QgsWkbTypes.geometryDisplayString(lyr.geometryType())
                        if hasattr(lyr, "geometryType") else "raster")
            rows.append([table_item(node.name(), bold=group), table_item(kind), table_item(src, grey=True),
                         table_item(dest_txt, bold=True, fill="FBE3C8" if dest == m.UNSORTED_GROUP else None),
                         table_item(m.THEMES[theme][0] if theme is not None else "-", grey=theme is None)])
        fill_table(self.moves_tbl, ["Layer / group", "Kind", "Now in", "Moves to", "Type of information"], rows)

        srows = []
        n_style = 0
        if self.opts["APPLY_STYLES"].isChecked():
            for layer in sorted(project.mapLayers().values(), key=lambda l: l.name().lower()):
                files, reason = m.styles_for(layer)
                if files:
                    n_style += 1
                    srows.append([table_item(layer.name()), table_item(", ".join(files))])
                elif reason:
                    srows.append([table_item(layer.name(), grey=True), table_item(reason, grey=True)])
        fill_table(self.styles_tbl, ["Layer", "Styles to apply"], srows)
        unsorted = sum(1 for _, d in moves if d == m.UNSORTED_GROUP)
        # dividers and Unsorted are rebuilt every run, so they are not real changes
        removed = [g for g in emptied if not m.is_separator(g.name()) and g.name() != m.UNSORTED_GROUP]
        self.statusBar().showMessage(
            f"{len(moves)} move(s), {len(removed)} group(s) to merge or remove, {n_style} layer(s) to style"
            + (f", {unsorted} to Unsorted" if unsorted else ""))

    # ---- run ----------------------------------------------------------------------
    def run(self):
        if self.running:
            return
        m = self.load()  # fresh copy of the script
        boundary = self.configure(m)
        self.save_settings()
        self.run_logged("Project Arrange", m.main, boundary=boundary, report_to=REPORT_TO)
        self.mod = m
        self.safe_refresh()
