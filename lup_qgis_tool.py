r"""
QGIS Processing tool: Land Use Plan layer -> Excel database + PowerPoint tables.

Install with Install_QGIS_Tool.bat: it copies this file to <profile>\processing\scripts
and lup_pipeline.py, the slide template and python-pptx to <profile>\lup_tool.
The tool then appears in the Processing Toolbox under
Scripts > Land Use Plan > "Export Land Use (Excel + PowerPoint)".

All calculations live in lup_pipeline.py; this file only reads the QGIS layer
and hands its attributes to that module.
"""

import os
import sys
import importlib

from qgis.PyQt.QtCore import QCoreApplication, QVariant, QDate, QDateTime
from qgis.core import (
    QgsDistanceArea,
    QgsProcessing,
    QgsProcessingAlgorithm,
    QgsProcessingException,
    QgsProcessingParameterBoolean,
    QgsProcessingParameterEnum,
    QgsProcessingParameterFeatureSource,
    QgsProcessingParameterField,
    QgsProcessingParameterFile,
    QgsProcessingParameterFileDestination,
    QgsProcessingParameterNumber,
    QgsProcessingParameterDefinition,
    QgsUnitTypes,
)

LIB_DIR_NAME = "lup_tool"  # inside the QGIS profile folder, NOT inside processing\scripts
TEMPLATE_NAME = "LUP_template.pptx"


def _lib_dirs():
    """Folders that hold lup_pipeline.py, best first.

    The library must not sit inside processing\\scripts: QGIS scans that folder
    recursively and would try to load every .py file (incl. python-pptx) as a script.
    """
    cands = [os.environ.get("LUP_TOOL_DIR", "")]
    try:
        cands.append(os.path.dirname(os.path.abspath(__file__)))  # plugin folder / next to lup_pipeline.py
    except NameError:  # run from the script editor
        pass
    cands.append(_profile_dir())
    out = []
    for c in cands:
        if c and os.path.exists(os.path.join(c, "lup_pipeline.py")) and c not in out:
            out.append(c)
    return out


def _profile_dir():
    try:
        from qgis.core import QgsApplication
        return os.path.join(QgsApplication.qgisSettingsDirPath(), LIB_DIR_NAME)
    except Exception:
        return ""


def _data_dir(code_dir=None):
    """Where colour settings, template and library copy live: <profile>/lup_tool when it
    exists (plugin / installer), else next to the code."""
    p = _profile_dir()
    if p and any(os.path.exists(os.path.join(p, f)) for f in ("lup_colours.json", TEMPLATE_NAME)):
        return p
    return code_dir or (_lib_dirs() or [""])[0]


def _load_pipeline():
    dirs = _lib_dirs()
    if not dirs:
        raise QgsProcessingException(
            "lup_pipeline.py not found. Run Install_QGIS_Tool.bat again "
            f"(it installs to <QGIS profile>\\{LIB_DIR_NAME}).")
    if dirs[0] not in sys.path:
        sys.path.insert(0, dirs[0])
    vendor = os.path.join(dirs[0], "vendor")  # python-pptx + XlsxWriter, installed by the .bat
    if os.path.isdir(vendor) and vendor not in sys.path:
        sys.path.append(vendor)  # appended so QGIS's own lxml/openpyxl still take priority
    try:
        import lup_pipeline
        return importlib.reload(lup_pipeline), dirs[0]  # pick up edits without restarting QGIS
    except ImportError as e:
        raise QgsProcessingException(
            f"Missing Python package ({e.name}). Run Install_QGIS_Tool.bat again "
            "(it needs internet once to download python-pptx).")


def _py(v):
    """QGIS attribute value -> plain Python value."""
    if v is None or (isinstance(v, QVariant) and v.isNull()):
        return None
    if isinstance(v, (QDate, QDateTime)):
        return v.toString("yyyy-MM-dd")
    return v


class LupTablesAlgorithm(QgsProcessingAlgorithm):
    INPUT = "INPUT"
    SITE_TOTAL = "SITE_TOTAL"
    AREA_SOURCE = "AREA_SOURCE"
    NUMBER_STYLE = "NUMBER_STYLE"
    TEMPLATE = "TEMPLATE"
    OPEN = "OPEN_OUTPUTS"
    COLOR_FIELD = "COLOR_FIELD"
    COLOR_PRIORITY = "COLOR_PRIORITY"
    LIBRARY = "LIBRARY"
    PIVOTS = "PIVOTS"
    PRIORITY_KEYS = ["layer", "library", "palette"]
    PRIORITY_OPTIONS = ["Layer symbology → colour library → project palette",
                        "Colour library → layer symbology → project palette",
                        "Project palette → colour library"]
    OUT_XLSX = "OUTPUT_XLSX"
    OUT_PPTX = "OUTPUT_PPTX"
    # attribute field parameters: (param name, label, default field, required)
    FIELDS = [
        ("FLD_TYPE", "Saleable / Non-Saleable field", "Type", True),
        ("FLD_SUBZ", "Land use (SUBZONA) field", "SUBZONA", True),
        ("FLD_AREA", "Area (sqm) field", "Area_sqm", False),
        ("FLD_INDT", "Industrial type field", "IND_TYPE", False),
        ("FLD_PHASE", "Phase field", "Phase", False),
        ("FLD_PARCEL", "Parcel ID field", "Parcel_ID", False),
        ("FLD_POP", "Population field", "Pop_Total", False),
        ("FLD_POPTYPE", "Population type field", "Pop_Type", False),
    ]
    AREA_OPTIONS = ["Area field (attribute table)", "Recalculate from polygon geometry"]
    STYLE_OPTIONS = ["Indonesian  (1.234.567 / 12,34)", "English  (1,234,567 / 12.34)"]

    def tr(self, s):
        return QCoreApplication.translate("LupTablesAlgorithm", s)

    def createInstance(self):
        return LupTablesAlgorithm()

    def name(self):
        return "lup_tables"

    def displayName(self):
        return self.tr("Export Land Use (Excel + PowerPoint)")

    def group(self):
        return self.tr("Land Use Plan")

    def groupId(self):
        return "landuseplan"

    def shortHelpString(self):
        return self.tr(
            "Builds the Land Use Plan Excel database and PowerPoint table/chart slides "
            "from a land-use polygon layer.\n\n"
            "Excel: Overall MP (Saleable / Non-Saleable summary, industrial breakdown, "
            "population), Phasing, Intensitas (BCR, FAR, GFA per parcel), Source, Checks. "
            "Summary sheets are SUMIFS formulas on the Source sheet.\n\n"
            "PowerPoint: land-use summary table with pie + Saleable ring, and program "
            "table with industrial breakdown chart.\n\n"
            "Site total: GRAND TOTAL area in sqm; Road absorbs the difference to the data "
            "total. Use 0 for the plain sum of the layer.\n\n"
            "Tick 'Selected features only' to report on a selection. Field names can be "
            "changed under Advanced Parameters.")

    def initAlgorithm(self, config=None):
        self.addParameter(QgsProcessingParameterFeatureSource(
            self.INPUT, self.tr("Land use layer"), [QgsProcessing.TypeVectorPolygon]))
        self.addParameter(QgsProcessingParameterNumber(
            self.SITE_TOTAL, self.tr("Site total area (sqm), 0 = sum of layer"),
            QgsProcessingParameterNumber.Double, defaultValue=5_000_000, minValue=0))
        self.addParameter(QgsProcessingParameterEnum(
            self.AREA_SOURCE, self.tr("Take parcel area from"),
            options=self.AREA_OPTIONS, defaultValue=0))
        self.addParameter(QgsProcessingParameterEnum(
            self.NUMBER_STYLE, self.tr("Number format on slides"),
            options=self.STYLE_OPTIONS, defaultValue=0))

        default_tpl = ""
        dirs = _lib_dirs()
        for d in ([_data_dir(dirs[0] if dirs else None)] + dirs[:1] if dirs else []):
            if d and os.path.exists(os.path.join(d, TEMPLATE_NAME)):
                default_tpl = os.path.join(d, TEMPLATE_NAME)
                break
        self.addParameter(QgsProcessingParameterFile(
            self.TEMPLATE, self.tr("PowerPoint template (optional)"),
            extension="pptx", defaultValue=default_tpl or None, optional=True))

        self.addParameter(QgsProcessingParameterField(
            self.COLOR_FIELD, self.tr("Colour by field (empty = layer style field, else SUBZONA)"),
            parentLayerParameterName=self.INPUT, optional=True))
        self.addParameter(QgsProcessingParameterEnum(
            self.COLOR_PRIORITY, self.tr("Colour source"), options=self.PRIORITY_OPTIONS, defaultValue=0))
        self.addParameter(QgsProcessingParameterFile(
            self.LIBRARY, self.tr("Colour library (empty = the one set in the Export Land Use window)"),
            extension="xlsx", optional=True))
        self.addParameter(QgsProcessingParameterBoolean(
            self.PIVOTS, self.tr("Summary sheets as Excel pivot tables (needs Microsoft Excel)"), defaultValue=True))

        for pname, label, default, required in self.FIELDS:
            p = QgsProcessingParameterField(
                pname, self.tr(label), defaultValue=default,
                parentLayerParameterName=self.INPUT, optional=not required)
            p.setFlags(p.flags() | QgsProcessingParameterDefinition.FlagAdvanced)
            self.addParameter(p)

        self.addParameter(QgsProcessingParameterBoolean(
            self.OPEN, self.tr("Open outputs when finished"), defaultValue=True))
        self.addParameter(QgsProcessingParameterFileDestination(
            self.OUT_XLSX, self.tr("Excel database"), "Excel files (*.xlsx)", optional=True))
        self.addParameter(QgsProcessingParameterFileDestination(
            self.OUT_PPTX, self.tr("PowerPoint tables"), "PowerPoint files (*.pptx)", optional=True))

    def processAlgorithm(self, parameters, context, feedback):
        L, lib_dir = _load_pipeline()
        feedback.pushInfo(f"Using {os.path.join(lib_dir, 'lup_pipeline.py')}")

        source = self.parameterAsSource(parameters, self.INPUT, context)
        if source is None:
            raise QgsProcessingException(self.invalidSourceError(parameters, self.INPUT))

        # map chosen field names onto the pipeline's field constants
        chosen = {}
        for pname, _, default, _ in self.FIELDS:
            chosen[pname] = self.parameterAsString(parameters, pname, context) or default
        L.F_TYPE, L.F_SUBZ, L.F_AREA = chosen["FLD_TYPE"], chosen["FLD_SUBZ"], chosen["FLD_AREA"]
        L.F_INDT, L.F_PHASE, L.F_PARCEL = chosen["FLD_INDT"], chosen["FLD_PHASE"], chosen["FLD_PARCEL"]
        L.F_POP, L.F_POPTYPE = chosen["FLD_POP"], chosen["FLD_POPTYPE"]

        use_geom = self.parameterAsEnum(parameters, self.AREA_SOURCE, context) == 1
        fields = [f.name() for f in source.fields()]
        if not use_geom and L.F_AREA not in fields:
            raise QgsProcessingException(
                f"Area field '{L.F_AREA}' not found. Pick another field or choose "
                "'Recalculate from polygon geometry'.")
        if use_geom and L.F_AREA not in fields:
            fields.append(L.F_AREA)

        # area measurement: planar in projected CRS, ellipsoidal if the layer is geographic
        da = QgsDistanceArea()
        da.setSourceCrs(source.sourceCrs(), context.transformContext())
        if source.sourceCrs().isGeographic():
            da.setEllipsoid(context.ellipsoid() or "EPSG:7030")
            feedback.pushWarning("Layer CRS is geographic; polygon areas are measured on the ellipsoid.")

        rows = []
        n = source.featureCount() or 1
        for i, feat in enumerate(source.getFeatures()):
            if feedback.isCanceled():
                return {}
            rec = {name: _py(v) for name, v in zip([f.name() for f in source.fields()], feat.attributes())}
            geom = feat.geometry()
            area = 0.0
            if geom and not geom.isEmpty():
                area = da.convertAreaMeasurement(da.measureArea(geom), QgsUnitTypes.AreaSquareMeters)
            rec["_geom_area"] = area
            if use_geom:
                rec[L.F_AREA] = round(area, 2)
            rows.append(rec)
            feedback.setProgress(int(60 * (i + 1) / n))
        if not rows:
            raise QgsProcessingException("The layer (or selection) has no features.")

        xlsx = self.parameterAsFileOutput(parameters, self.OUT_XLSX, context)
        pptx = self.parameterAsFileOutput(parameters, self.OUT_PPTX, context)
        if not xlsx and not pptx:
            raise QgsProcessingException("Choose at least one output (Excel or PowerPoint).")
        template = self.parameterAsFile(parameters, self.TEMPLATE, context) or None
        site_total = self.parameterAsDouble(parameters, self.SITE_TOTAL, context)
        style = ["id", "en"][self.parameterAsEnum(parameters, self.NUMBER_STYLE, context)]

        # colours: layer symbology (if INPUT is a layer) -> library -> palette
        layer = None
        try:
            layer = self.parameterAsVectorLayer(parameters, self.INPUT, context)
        except Exception:
            pass
        import lup_colors as LC
        cfield = self.parameterAsString(parameters, self.COLOR_FIELD, context)
        if not cfield:
            cfield = (LC.renderer_field(layer) if layer else None) or L.F_SUBZ
        colors = L.make_colors(fields, rows, cfield,
                               self.PRIORITY_KEYS[self.parameterAsEnum(parameters, self.COLOR_PRIORITY, context)],
                               self.parameterAsFile(parameters, self.LIBRARY, context) or None,
                               LC.layer_colors(layer, cfield) if layer else {}, _data_dir(lib_dir))
        try:
            S, issues = L.run(fields, rows, xlsx, pptx, site_total, template, style,
                              log=feedback.pushInfo, colors=colors,
                              pivots=self.parameterAsBoolean(parameters, self.PIVOTS, context))
        except (RuntimeError, ValueError) as e:
            raise QgsProcessingException(str(e))
        feedback.setProgress(100)

        for k, v in issues[4:]:
            if v != "OK":
                feedback.pushWarning(f"{k}: {v}")
        feedback.pushInfo("")
        for line in L.summary_lines(S):
            feedback.pushInfo(line)

        if self.parameterAsBoolean(parameters, self.OPEN, context):
            for p in (xlsx, pptx):
                if p and os.path.exists(p):
                    try:
                        os.startfile(p)  # Windows
                    except (AttributeError, OSError):
                        from qgis.PyQt.QtCore import QUrl
                        from qgis.PyQt.QtGui import QDesktopServices
                        QDesktopServices.openUrl(QUrl.fromLocalFile(p))
        return {self.OUT_XLSX: xlsx, self.OUT_PPTX: pptx}
