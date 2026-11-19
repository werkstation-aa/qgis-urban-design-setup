"""QGIS plugin: Land Use Population.

Toolbar button, Vector menu entry and layer right-click entry open the Land Use Population
window for the active polygon layer.
"""

import os

from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtWidgets import QAction
from qgis.core import QgsMapLayer, QgsWkbTypes

import sys

PLUGIN_DIR = os.path.dirname(os.path.abspath(__file__))
VENDOR = os.path.join(PLUGIN_DIR, "vendor")
if os.path.isdir(VENDOR) and VENDOR not in sys.path:
    sys.path.append(VENDOR)  # python-pptx, XlsxWriter; appended so QGIS's own packages keep priority
ICON = os.path.join(PLUGIN_DIR, "icon.svg")
MENU = "&Land Use Population"


class LandUsePopulationPlugin:
    def __init__(self, iface):
        self.iface = iface
        self.action = None
        self.layer_action = None
        self.window = None

    def initGui(self):
        win = self.iface.mainWindow()
        self.action = QAction(QIcon(ICON), "Land Use Population…", win)
        self.action.setToolTip("Calculate resident / daytime population of a land-use plan (SNI 03-1733, PU 26/2008)")
        self.action.triggered.connect(lambda: self.run(self.iface.activeLayer()))
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToVectorMenu(MENU, self.action)
        self.layer_action = QAction(QIcon(ICON), "Calculate Population…", win)
        self.layer_action.triggered.connect(lambda: self.run(self.iface.activeLayer()))
        self.iface.addCustomActionForLayerType(self.layer_action, "", QgsMapLayer.VectorLayer, True)

    def unload(self):
        if self.window is not None:
            self.window.close()
            self.window = None
        if self.layer_action is not None:
            self.iface.removeCustomActionForLayerType(self.layer_action)
        if self.action is not None:
            self.iface.removePluginVectorMenu(MENU, self.action)
            self.iface.removeToolBarIcon(self.action)

    def run(self, layer=None):
        from .pop_widget import PopulationWindow
        if self.window is None or not self.window.isVisible():
            self.window = PopulationWindow(self.iface, self.iface.mainWindow())
        if (layer is not None and layer.type() == QgsMapLayer.VectorLayer
                and layer.geometryType() == QgsWkbTypes.PolygonGeometry):
            self.window.set_layer(layer)
        self.window.show()
        self.window.showNormal()
        self.window.raise_()
        self.window.activateWindow()
