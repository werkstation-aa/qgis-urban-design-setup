"""QGIS plugin: Export Land Use window (toolbar button / menu) + Processing tool.

Self-contained: the calculations (lup_pipeline.py), window (lup_widget.py), colour
scheme (lup_colors.py), Processing algorithm (lup_qgis_tool.py) and python-pptx /
comtypes (vendor/) ship inside this plugin folder.

User data (colour settings, template, colour library copy) lives in
<QGIS profile>/lup_tool so it survives plugin updates.
"""

import importlib
import os
import shutil
import sys

from qgis.PyQt.QtCore import Qt
from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtWidgets import QAction, QMessageBox
from qgis.core import QgsApplication, QgsProcessingProvider

MENU = "&Land Use Plan"
PLUGIN_DIR = os.path.dirname(os.path.abspath(__file__))
MODULES = ("lup_pipeline", "lup_colors", "lup_widget", "lup_qgis_tool")
SEED = ("lup_colours.json", "LUP_template.pptx", "RDTR_Colour Code.xlsx")


def data_dir():
    d = os.path.join(QgsApplication.qgisSettingsDirPath(), "lup_tool")
    os.makedirs(d, exist_ok=True)
    for f in SEED:  # first run: copy defaults; later runs keep the user's versions
        src, dst = os.path.join(PLUGIN_DIR, "data", f), os.path.join(d, f)
        if os.path.exists(src) and not os.path.exists(dst):
            shutil.copy2(src, dst)
    return d


def load_modules(reload=False):
    """Import the bundled modules from this plugin folder (not an older copy elsewhere)."""
    for p in (os.path.join(PLUGIN_DIR, "vendor"), PLUGIN_DIR):
        if p in sys.path:
            sys.path.remove(p)
    sys.path.insert(0, PLUGIN_DIR)
    sys.path.append(os.path.join(PLUGIN_DIR, "vendor"))  # last: QGIS's own lxml/openpyxl win
    for name in MODULES:
        mod = sys.modules.get(name)
        if mod is not None and (reload or os.path.dirname(os.path.abspath(mod.__file__)) != PLUGIN_DIR):
            del sys.modules[name]
    return {name: importlib.import_module(name) for name in MODULES}


class LupProvider(QgsProcessingProvider):
    def id(self):
        return "lup"

    def name(self):
        return "Land Use Plan"

    def icon(self):
        return QIcon(os.path.join(PLUGIN_DIR, "icon.svg"))

    def loadAlgorithms(self):
        try:
            tool = load_modules()["lup_qgis_tool"]
            self.addAlgorithm(tool.LupTablesAlgorithm())
        except ImportError as e:
            QgsApplication.messageLog().logMessage(f"Export Land Use: could not load ({e})", "Export Land Use")


class LupTablesPlugin:
    def __init__(self, iface):
        self.iface = iface
        self.action = None
        self.window = None
        self.provider = None

    def initProcessing(self):
        self.provider = LupProvider()
        QgsApplication.processingRegistry().addProvider(self.provider)

    def initGui(self):
        self.initProcessing()
        icon = QIcon(os.path.join(PLUGIN_DIR, "icon.svg"))
        self.action = QAction(icon, "Export Land Use (Excel + PowerPoint)…", self.iface.mainWindow())
        self.action.setToolTip("Preview and export Land Use Plan tables and charts")
        self.action.triggered.connect(self.run)
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu(MENU, self.action)

    def unload(self):
        if self.window is not None:
            self.window.close()
            self.window = None
        if self.provider is not None:
            QgsApplication.processingRegistry().removeProvider(self.provider)
            self.provider = None
        self.iface.removePluginMenu(MENU, self.action)
        self.iface.removeToolBarIcon(self.action)

    def run(self):
        if self.window is not None and self.window.isVisible():
            self.window.showNormal()
            self.window.raise_()
            self.window.activateWindow()
            return
        try:
            mods = load_modules(reload=True)  # picks up edits without restarting QGIS
        except ImportError as e:
            QMessageBox.critical(self.iface.mainWindow(), "Export Land Use",
                                 f"Missing Python package: {e.name}\n\nReinstall the plugin zip.")
            return
        self.window = mods["lup_widget"].LupWindow(self.iface, self.iface.mainWindow(), lib_dir=data_dir())
        self.window.setWindowFlags(self.window.windowFlags() | Qt.Window)
        layer = self.iface.activeLayer()
        if layer is not None and hasattr(layer, "geometryType"):
            try:
                self.window.layer_cb.setLayer(layer)
            except Exception:
                pass
        self.window.show()
