"""Urban Design Setup - every tool of the Automation folder in one plugin
(plugin folder / Processing id keep the original names 'qgis_automation_pack' / 'automation').

Every tool opens its own window; nothing runs in the Python Console.

Layout inside the installed plugin folder (assembled by build_pack_zip.py):
    plugin.py           menu, toolbar, Processing provider, opens the windows
    tool_window.py      shared window frame (settings | preview tabs | run log)
    setup_window.py     Project Setup window     -> scripts/qgis_project_setup_r2.py
    arrange_window.py   Project Arrange window   -> scripts/qgis_project_arrange.py
    scripts/            the console scripts, imported (never run as scripts) by the windows;
                        Layout Automation opens its own dialog from qgis_layout_automation.py
    lup/                Export Land Use: window + Processing algorithm + calculations
    vector_pivot/       Export Pivot: preview window + Processing algorithm + Excel writer
    population/         Land Use Population: assumption preview window, memory layer, Excel + PowerPoint
    vendor/             python-pptx, XlsxWriter, typing_extensions, comtypes
    data/               LUP defaults, copied to <QGIS profile>/lup_tool on first run
"""

import contextlib
import importlib
import os
import shutil
import sys
import traceback

from qgis.PyQt.QtCore import Qt
from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtWidgets import QAction, QMessageBox
from qgis.core import QgsApplication, QgsMapLayer, QgsMessageLog, QgsProcessingProvider, Qgis

PLUGIN_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(PLUGIN_DIR, "scripts")
LUP_DIR = os.path.join(PLUGIN_DIR, "lup")
VENDOR = os.path.join(PLUGIN_DIR, "vendor")
TITLE = "Urban Design Setup"
MENU = "&" + TITLE
LOG = TITLE
PROVIDER_ID = "automation"

SETUP_SCRIPT = os.path.join(SCRIPTS, "qgis_project_setup_r2.py")
STYLES_DIR = os.path.join(SCRIPTS, "styles")  # QML styles shared by Project Setup and Project Arrange
ARRANGE_SCRIPT = os.path.join(SCRIPTS, "qgis_project_arrange.py")
LAYOUT_SCRIPT = os.path.join(SCRIPTS, "qgis_layout_automation.py")
LUP_MODULES = ("lup_pipeline", "lup_colors", "lup_widget", "lup_qgis_tool")
LUP_SEED = ("lup_colours.json", "LUP_template.pptx", "RDTR_Colour Code.xlsx")


def setup_paths():
    if LUP_DIR in sys.path:
        sys.path.remove(LUP_DIR)
    sys.path.insert(0, LUP_DIR)
    if os.path.isdir(VENDOR) and VENDOR not in sys.path:
        sys.path.append(VENDOR)  # last: QGIS's own lxml / openpyxl keep priority


def icon(name):
    return QIcon(os.path.join(PLUGIN_DIR, "icons", name))


def lup_data_dir():
    d = os.path.join(QgsApplication.qgisSettingsDirPath(), "lup_tool")
    os.makedirs(d, exist_ok=True)
    for f in LUP_SEED:  # first run: copy defaults; later runs keep the user's versions
        src, dst = os.path.join(PLUGIN_DIR, "data", f), os.path.join(d, f)
        if os.path.exists(src) and not os.path.exists(dst):
            shutil.copy2(src, dst)
    return d


def load_lup(reload=False):
    """Import the LUP modules from lup/ (not an older copy from another plugin)."""
    setup_paths()
    for name in LUP_MODULES:
        mod = sys.modules.get(name)
        if mod is not None and (reload or os.path.dirname(os.path.abspath(mod.__file__)) != LUP_DIR):
            del sys.modules[name]
    return {name: importlib.import_module(name) for name in LUP_MODULES}


class AutomationProvider(QgsProcessingProvider):
    def id(self):
        return PROVIDER_ID

    def name(self):
        return TITLE

    def icon(self):
        return icon("pack.svg")

    def loadAlgorithms(self):
        try:
            self.addAlgorithm(load_lup()["lup_qgis_tool"].LupTablesAlgorithm())
        except Exception as e:
            QgsMessageLog.logMessage(f"Export Land Use not loaded: {e}", LOG, Qgis.Warning)
        try:
            from .vector_pivot.vpe_algorithm import VectorPivotAlgorithm
            self.addAlgorithm(VectorPivotAlgorithm())
        except Exception as e:
            QgsMessageLog.logMessage(f"Export Pivot not loaded: {e}", LOG, Qgis.Warning)


class AutomationPackPlugin:
    def __init__(self, iface):
        self.iface = iface
        self.provider = None
        self.toolbar = None
        self.actions = []
        self.layer_action = None
        self.pop_layer_action = None
        self.windows = {}

    # ---- setup -----------------------------------------------------------
    def initProcessing(self):
        setup_paths()
        self.provider = AutomationProvider()
        QgsApplication.processingRegistry().addProvider(self.provider)

    def initGui(self):
        self.initProcessing()
        self.toolbar = self.iface.addToolBar(TITLE)
        self.toolbar.setObjectName("AutomationPackToolbar")
        self._add("Project Setup…", "Build the standard layer tree, basemaps, live datasets and OSM roads",
                  "setup.svg", self.open_setup)
        self._add("Project Arrange…", "Sort the loaded layers into the project setup tree and style them",
                  "arrange.svg", self.open_arrange)
        self._add("Layout Automation…", "Bookmark x map themes -> print layout pages and 300 dpi images",
                  "layout.svg", self.open_layout)
        self._add("Export Land Use (Excel + PowerPoint)…", "Preview and export Land Use Plan tables and charts",
                  "lup.svg", self.open_lup, separator=True)
        self._add("Export Pivot to Excel…", "Preview a pivot table of the active vector layer, coloured by its "
                  "symbology, and export it to Excel", "pivot.svg", lambda: self.open_pivot(self.iface.activeLayer()))
        self._add("Land Use Population…", "Resident / visitor / worker population of a land-use plan from verified "
                  "standards (SNI 03-1733, PU 26/2008), with assumption preview, Excel and PowerPoint",
                  "population.svg", lambda: self.open_population(self.iface.activeLayer()))
        self.layer_action = QAction(icon("pivot.svg"), "Export Pivot to Excel…", self.iface.mainWindow())
        self.layer_action.triggered.connect(lambda: self.open_pivot(self.iface.activeLayer()))
        self.iface.addCustomActionForLayerType(self.layer_action, "", QgsMapLayer.VectorLayer, True)
        self.pop_layer_action = QAction(icon("population.svg"), "Calculate Population…", self.iface.mainWindow())
        self.pop_layer_action.triggered.connect(
            lambda: self.guarded("Land Use Population", lambda: self.open_population(self.iface.activeLayer())))
        self.iface.addCustomActionForLayerType(self.pop_layer_action, "", QgsMapLayer.VectorLayer, True)

    def _add(self, text, tip, ico, slot, separator=False):
        a = QAction(icon(ico), text, self.iface.mainWindow())
        a.setToolTip(tip)
        a.setStatusTip(tip)
        a.triggered.connect(lambda _=False: self.guarded(text, slot))
        if separator:
            self.toolbar.addSeparator()
        self.toolbar.addAction(a)
        self.iface.addPluginToMenu(MENU, a)
        self.actions.append(a)
        return a

    def unload(self):
        for w in self.windows.values():
            try:
                w.close()
            except RuntimeError:
                pass
        self.windows = {}
        if self.layer_action is not None:
            self.iface.removeCustomActionForLayerType(self.layer_action)
        if self.pop_layer_action is not None:
            self.iface.removeCustomActionForLayerType(self.pop_layer_action)
        for a in self.actions:
            self.iface.removePluginMenu(MENU, a)
        self.actions = []
        if self.toolbar is not None:
            self.toolbar.deleteLater()
            self.toolbar = None
        if self.provider is not None:
            QgsApplication.processingRegistry().removeProvider(self.provider)
            self.provider = None

    # ---- helpers ---------------------------------------------------------
    def guarded(self, title, fn):
        """Open a tool; a failure is reported in a message box and the log, not the console."""
        try:
            fn()
        except Exception:
            tb = traceback.format_exc()
            QgsMessageLog.logMessage(tb, LOG, Qgis.Critical)
            QMessageBox.critical(self.iface.mainWindow(), title.rstrip("…"),
                                 "Could not open the tool:\n\n" + tb.strip().splitlines()[-1]
                                 + f"\n\nDetails: View > Panels > Log Messages > {LOG}.")

    def show_window(self, key, factory):
        """One window per tool: bring it back if open, else build it."""
        w = self.windows.get(key)
        if w is not None:
            try:
                if w.isVisible():
                    w.showNormal()
                    w.raise_()
                    w.activateWindow()
                    return w
            except RuntimeError:  # deleted by Qt
                pass
        w = factory()
        self.windows[key] = w
        w.show()
        return w

    # ---- tools -----------------------------------------------------------
    def open_setup(self):
        from .setup_window import SetupWindow
        self.show_window("setup", lambda: SetupWindow(self.iface, SETUP_SCRIPT, self.iface.mainWindow()))

    def open_arrange(self):
        from .arrange_window import ArrangeWindow
        self.show_window("arrange", lambda: ArrangeWindow(self.iface, ARRANGE_SCRIPT, STYLES_DIR,
                                                          self.iface.mainWindow()))

    def open_layout(self):
        """Layout Automation has its own dialog: import the script and open it."""
        from .tool_window import LogStream, load_script
        mod = load_script(LAYOUT_SCRIPT, "qgis_layout_automation")
        stream = LogStream(lambda line: QgsMessageLog.logMessage(line, LOG, Qgis.Info))
        with contextlib.redirect_stdout(stream):  # its report goes to the Log Messages panel
            mod.run()
        stream.flush()

    def open_lup(self):
        w = self.windows.get("lup")
        if w is not None:
            try:
                if w.isVisible():
                    w.showNormal()
                    w.raise_()
                    w.activateWindow()
                    return
            except RuntimeError:
                pass
        try:
            mods = load_lup(reload=True)
        except ImportError as e:
            QMessageBox.critical(self.iface.mainWindow(), "Export Land Use",
                                 f"Missing Python package: {e.name}\n\nReinstall the plugin zip.")
            return
        w = mods["lup_widget"].LupWindow(self.iface, self.iface.mainWindow(), lib_dir=lup_data_dir())
        w.setWindowFlags(w.windowFlags() | Qt.Window)
        layer = self.iface.activeLayer()
        if layer is not None and hasattr(layer, "geometryType"):
            try:
                w.layer_cb.setLayer(layer)
            except Exception:
                pass
        self.windows["lup"] = w
        w.show()

    def open_population(self, layer=None):
        from qgis.core import QgsWkbTypes
        from .population.pop_widget import PopulationWindow
        w = self.show_window("population", lambda: PopulationWindow(self.iface, self.iface.mainWindow()))
        if (layer is not None and layer.type() == QgsMapLayer.VectorLayer
                and layer.geometryType() == QgsWkbTypes.PolygonGeometry):
            w.set_layer(layer)

    def open_pivot(self, layer=None):
        from .vector_pivot.vpe_widget import PivotWindow
        w = self.show_window("pivot", lambda: PivotWindow(self.iface, self.iface.mainWindow()))
        if layer is not None and layer.type() == QgsMapLayer.VectorLayer:
            w.set_layer(layer)
            w.sel_only.setChecked(layer.selectedFeatureCount() > 0)
