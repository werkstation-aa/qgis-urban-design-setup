"""Processing algorithm: vector layer -> Excel workbook with a pivot table coloured by
the layer symbology. Reading the layer lives in vpe_layer.py (shared with the preview
window), writing Excel in vpe_core.py (no QGIS imports)."""

import os

from qgis.PyQt.QtCore import QCoreApplication
from qgis.core import (
    QgsProcessing,
    QgsProcessingAlgorithm,
    QgsProcessingException,
    QgsProcessingParameterBoolean,
    QgsProcessingParameterEnum,
    QgsProcessingParameterField,
    QgsProcessingParameterFileDestination,
    QgsProcessingParameterVectorLayer,
)

from . import vpe_core as core
from . import vpe_layer as VL


def open_file(path):
    try:
        os.startfile(path)  # Windows
    except (AttributeError, OSError):
        from qgis.PyQt.QtCore import QUrl
        from qgis.PyQt.QtGui import QDesktopServices
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))


class VectorPivotAlgorithm(QgsProcessingAlgorithm):
    INPUT, SELECTED = "INPUT", "SELECTED_ONLY"
    ROWS, COLUMN, VALUES = "ROW_FIELDS", "COLUMN_FIELD", "VALUE_FIELDS"
    AGG, MEASURE, COUNT, PERCENT = "AGGREGATE", "MEASURE", "ADD_COUNT", "PERCENT"
    PIVOT, OPEN, OUTPUT = "PIVOT_TYPE", "OPEN_OUTPUT", "OUTPUT"

    def tr(self, s):
        return QCoreApplication.translate("VectorPivotAlgorithm", s)

    def createInstance(self):
        return VectorPivotAlgorithm()

    def name(self):
        return "export"

    def displayName(self):
        return self.tr("Export pivot table to Excel (coloured by symbology)")

    def group(self):
        return ""

    def groupId(self):
        return ""

    def flags(self):
        # main thread: the layer renderer is read directly and Excel is driven over COM
        try:
            from qgis.core import Qgis
            return super().flags() | Qgis.ProcessingAlgorithmFlag.NoThreading  # QGIS 3.36+
        except AttributeError:
            return super().flags() | QgsProcessingAlgorithm.FlagNoThreading

    def shortHelpString(self):
        return self.tr(
            "Exports a vector layer's attribute table to an .xlsx workbook with a pivot table "
            "already built and colour coded by the layer's symbology.\n\n"
            "Sheets: <b>Pivot</b> (rows / column / values you choose), <b>Source</b> (the attribute "
            "table as an Excel table, colour coded), <b>Colours</b> (legend).\n\n"
            "<b>Row fields</b>: left empty, the field the layer is categorised / graduated on is used, "
            "or the symbology class. A <i>Symbology</i> column (legend label of each feature) is always "
            "added to Source.\n\n"
            "<b>Colours</b>: every value of the row and column fields gets the colour of the features "
            "carrying it (the most common one when they differ), so it works with categorised, "
            "graduated, rule-based and single-symbol styles.\n\n"
            "With Microsoft Excel installed (Windows) the Pivot sheet is a real pivot table on the "
            "Source table: edit Source and use Data > Refresh All. Without Excel a static table "
            "with the same layout is written.\n\n"
            "For a live preview, use the Export Pivot to Excel window (toolbar or layer right-click).")

    def initAlgorithm(self, config=None):
        self.addParameter(QgsProcessingParameterVectorLayer(
            self.INPUT, self.tr("Vector layer"), [QgsProcessing.TypeVector]))
        self.addParameter(QgsProcessingParameterBoolean(
            self.SELECTED, self.tr("Selected features only"), defaultValue=False))
        self.addParameter(QgsProcessingParameterField(
            self.ROWS, self.tr("Row fields (in order; empty = symbology field)"),
            parentLayerParameterName=self.INPUT, allowMultiple=True, optional=True))
        self.addParameter(QgsProcessingParameterField(
            self.COLUMN, self.tr("Column field (optional)"),
            parentLayerParameterName=self.INPUT, optional=True))
        self.addParameter(QgsProcessingParameterField(
            self.VALUES, self.tr("Value fields (numeric, optional)"),
            parentLayerParameterName=self.INPUT, type=QgsProcessingParameterField.Numeric,
            allowMultiple=True, optional=True))
        self.addParameter(QgsProcessingParameterEnum(
            self.AGG, self.tr("Aggregate for value fields"), options=VL.AGG_OPTIONS, defaultValue=0))
        self.addParameter(QgsProcessingParameterEnum(
            self.MEASURE, self.tr("Add geometry measure"), options=[m[0] for m in VL.MEASURES], defaultValue=0))
        self.addParameter(QgsProcessingParameterBoolean(
            self.COUNT, self.tr("Add feature count"), defaultValue=True))
        self.addParameter(QgsProcessingParameterBoolean(
            self.PERCENT, self.tr("Add % of total for sums and counts"), defaultValue=True))
        self.addParameter(QgsProcessingParameterEnum(
            self.PIVOT, self.tr("Pivot sheet"), options=VL.PIVOT_OPTIONS, defaultValue=0))
        self.addParameter(QgsProcessingParameterBoolean(
            self.OPEN, self.tr("Open the workbook when done"), defaultValue=True))
        self.addParameter(QgsProcessingParameterFileDestination(
            self.OUTPUT, self.tr("Excel workbook"), "Excel workbook (*.xlsx)"))

    def processAlgorithm(self, parameters, context, feedback):
        layer = self.parameterAsVectorLayer(parameters, self.INPUT, context)
        if layer is None:
            raise QgsProcessingException(self.tr("Invalid input layer"))
        as_fields = getattr(self, "parameterAsStrings", None) or self.parameterAsFields  # 3.32+ / older
        out = self.parameterAsFileOutput(parameters, self.OUTPUT, context)
        if not out.lower().endswith(".xlsx"):
            out = os.path.splitext(out)[0] + ".xlsx"
        try:
            data = VL.collect(
                layer,
                selected=self.parameterAsBool(parameters, self.SELECTED, context),
                row_fields=as_fields(parameters, self.ROWS, context),
                col_field=(as_fields(parameters, self.COLUMN, context) or [None])[0],
                value_fields=as_fields(parameters, self.VALUES, context),
                agg=VL.AGGS[self.parameterAsEnum(parameters, self.AGG, context)],
                measure_index=self.parameterAsEnum(parameters, self.MEASURE, context),
                add_count=self.parameterAsBool(parameters, self.COUNT, context),
                percent=self.parameterAsBool(parameters, self.PERCENT, context),
                ellipsoid=context.ellipsoid(), transform_context=context.transformContext(),
                feedback=feedback)
        except VL.CollectError as e:
            raise QgsProcessingException(str(e))
        for w in data["warnings"]:
            feedback.pushWarning(w)
        if not data["rows"]:
            raise QgsProcessingException(self.tr("The layer has no features to export."))
        feedback.pushInfo(f"Row field(s): {', '.join(data['spec']['row_fields'])}")
        feedback.pushInfo(f"Writing {len(data['rows'])} features to {out}")
        pivot = VL.PIVOT_KEYS[self.parameterAsEnum(parameters, self.PIVOT, context)]
        try:
            made = core.write_workbook(out, data["fields"], data["rows"], data["spec"], data["colors"],
                                       data["orders"], native=pivot, title=data["title"], note=data["note"],
                                       log=feedback.pushInfo)
        except PermissionError:
            raise QgsProcessingException(self.tr(f"Cannot write {out} - is it open in Excel?"))
        except Exception as e:
            raise QgsProcessingException(f"{type(e).__name__}: {e}")
        feedback.setProgress(100)
        feedback.pushInfo("Pivot sheet: " + ("Excel pivot table" if made == "excel" else "static table"))
        if self.parameterAsBool(parameters, self.OPEN, context):
            open_file(out)
        return {self.OUTPUT: out}
