"""QGIS plugin: Vector Pivot Export.

Adds a Processing tool (Vector Pivot Export > Export pivot table to Excel), a toolbar
button / Vector menu entry and a layer right-click entry that open it for the layer.
XlsxWriter and comtypes ship in vendor/ inside this folder.
"""

import os
import sys

from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtWidgets import QAction
from qgis.core import QgsApplication, QgsMapLayer, QgsProcessingProvider

PLUGIN_DIR = os.path.dirname(os.path.abspath(__file__))
VENDOR = os.path.join(PLUGIN_DIR, "vendor")
if os.path.isdir(VENDOR) and VENDOR not in sys.path:
    sys.path.append(VENDOR)  # appended: packages QGIS already has keep priority

ICON = os.path.join(PLUGIN_DIR, "icon.svg")
ALG_ID = "vectorpivot:export"
MENU = "&Vector Pivot Export"


class VectorPivotProvider(QgsProcessingProvider):
    def id(self):
        return "vectorpivot"

    def name(self):
        return "Vector Pivot Export"

    def icon(self):
        return QIcon(ICON)

    def loadAlgorithms(self):
        try:
            from .vpe_algorithm import VectorPivotAlgorithm
            self.addAlgorithm(VectorPivotAlgorithm())
        except ImportError as e:
            QgsApplication.messageLog().logMessage(f"could not load ({e})", "Vector Pivot Export")


class VectorPivotPlugin:
    def __init__(self, iface):
        self.iface = iface
        self.provider = None
        self.action = None
        self.layer_action = None

    def initProcessing(self):
        self.provider = VectorPivotProvider()
        QgsApplication.processingRegistry().addProvider(self.provider)

    def initGui(self):
        self.initProcessing()
        win = self.iface.mainWindow()
        self.action = QAction(QIcon(ICON), "Export Pivot to Excel…", win)
        self.action.setToolTip("Export the active vector layer to Excel as a pivot table coloured by its symbology")
        self.action.triggered.connect(lambda: self.run(self.iface.activeLayer()))
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToVectorMenu(MENU, self.action)
        self.layer_action = QAction(QIcon(ICON), "Export Pivot to Excel…", win)
        self.layer_action.triggered.connect(lambda: self.run(self.iface.activeLayer()))
        self.iface.addCustomActionForLayerType(self.layer_action, "", QgsMapLayer.VectorLayer, True)

    def unload(self):
        if getattr(self, "window", None) is not None:
            self.window.close()
            self.window = None
        if self.layer_action is not None:
            self.iface.removeCustomActionForLayerType(self.layer_action)
        if self.action is not None:
            self.iface.removePluginVectorMenu(MENU, self.action)
            self.iface.removeToolBarIcon(self.action)
        if self.provider is not None:
            QgsApplication.processingRegistry().removeProvider(self.provider)
            self.provider = None

    def run(self, layer=None):
        """Open the preview window on the layer (the Processing tool stays in the Toolbox)."""
        from .vpe_widget import PivotWindow
        w = getattr(self, "window", None)
        if w is None or not w.isVisible():
            self.window = w = PivotWindow(self.iface, self.iface.mainWindow())
        if layer is not None and layer.type() == QgsMapLayer.VectorLayer:
            w.set_layer(layer)
            w.sel_only.setChecked(layer.selectedFeatureCount() > 0)
        w.show()
        w.showNormal()
        w.raise_()
        w.activateWindow()
