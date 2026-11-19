"""Shared window frame for the Urban Design Setup tools, laid out like the
Export Land Use window: settings on the left, preview tabs on the right, a run
log underneath. Console scripts are imported (not run) and driven from here;
whatever they print goes to the log instead of the Python Console."""

import contextlib
import importlib.util
import io
import traceback

from qgis.PyQt.QtCore import QEventLoop, QSettings, Qt, QTimer
from qgis.PyQt.QtGui import QColor, QKeySequence
from qgis.PyQt.QtWidgets import (QApplication, QComboBox, QFrame, QMainWindow, QPlainTextEdit,
                                 QPushButton, QScrollArea, QShortcut, QSizePolicy, QSplitter,
                                 QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget)

SETTINGS_ORG = "UrbanDesignSetup"


def load_script(path, module_name):
    """A fresh module from a console script, without running it: the scripts only
    start themselves when their __name__ is not module_name."""
    spec = importlib.util.spec_from_file_location(module_name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class LogStream(io.TextIOBase):
    """stdout replacement: complete lines go to a callback."""

    def __init__(self, emit):
        super().__init__()
        self.emit = emit
        self.buf = ""

    def write(self, s):
        self.buf += s
        while "\n" in self.buf:
            line, self.buf = self.buf.split("\n", 1)
            self.emit(line)
        return len(s)

    def flush(self):
        if self.buf:
            self.emit(self.buf)
            self.buf = ""


def make_table():
    t = QTableWidget()
    t.setEditTriggers(QTableWidget.NoEditTriggers)
    t.verticalHeader().setVisible(False)
    t.setSelectionBehavior(QTableWidget.SelectRows)
    t.setWordWrap(False)
    t.verticalHeader().setDefaultSectionSize(t.fontMetrics().height() + 8)
    return t


def is_dark(hx):
    r, g, b = int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) < 140


def table_item(text, fill=None, bold=False, right=False, grey=False):
    it = QTableWidgetItem("" if text is None else str(text))
    if fill:
        it.setBackground(QColor("#" + fill))
        it.setForeground(QColor("#FFFFFF" if is_dark(fill) else "#000000"))
    elif grey:
        it.setForeground(QColor("#888888"))
    if bold:
        f = it.font()
        f.setBold(True)
        it.setFont(f)
    if right:
        it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
    return it


def fill_table(t, headers, rows):
    """rows: list of lists of QTableWidgetItem (or plain values)."""
    t.clear()
    t.setColumnCount(len(headers))
    t.setRowCount(len(rows))
    t.setHorizontalHeaderLabels(headers)
    for i, cells in enumerate(rows):
        for j, c in enumerate(cells):
            t.setItem(i, j, c if isinstance(c, QTableWidgetItem) else table_item(c))
    t.resizeColumnsToContents()
    for j in range(len(headers)):
        t.setColumnWidth(j, max(60, min(360, t.columnWidth(j) + 14)))


def fit_panel(host):
    """Let combo boxes shrink so the settings panel never runs wider than its column."""
    for cb in host.findChildren(QComboBox):
        cb.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        cb.setMinimumContentsLength(8)
        cb.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)


class ToolWindow(QMainWindow):
    """Subclasses: set SETTINGS_KEY, call build_frame() in __init__, put group boxes
    in self.side, tabs in self.tabs, and implement refresh() (the preview)."""

    SETTINGS_KEY = "tool"

    def __init__(self, iface, parent=None, title=""):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() | Qt.Window)
        self.iface = iface
        self.settings = QSettings(SETTINGS_ORG, self.SETTINGS_KEY)
        self.setWindowTitle(title)
        self.setMinimumSize(640, 420)
        self.running = False
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.setInterval(350)
        self.timer.timeout.connect(self.safe_refresh)

    # ---- frame -------------------------------------------------------------
    def build_frame(self, run_text):
        form_host = QWidget()
        self.side = QVBoxLayout(form_host)
        left = QScrollArea()
        left.setWidgetResizable(True)
        left.setWidget(form_host)
        left.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        left.setMinimumWidth(280)
        left.setFrameShape(QFrame.NoFrame)
        self.run_btn = QPushButton(run_text + "  (Ctrl+Enter)")
        self.run_btn.setMinimumHeight(34)
        self.run_btn.setStyleSheet("font-weight:bold")
        self.run_btn.clicked.connect(self.run)
        QShortcut(QKeySequence("Ctrl+Return"), self, activated=self.run)
        QShortcut(QKeySequence("F5"), self, activated=self.safe_refresh)

        self.tabs = QTabWidget()
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(5000)
        self.vsplit = QSplitter(Qt.Vertical)
        self.vsplit.addWidget(self.tabs)
        self.vsplit.addWidget(self.log)
        self.vsplit.setStretchFactor(0, 5)
        self.vsplit.setStretchFactor(1, 1)
        self.vsplit.setSizes([560, 160])
        self.hsplit = QSplitter(Qt.Horizontal)
        self.hsplit.addWidget(left)
        self.hsplit.addWidget(self.vsplit)
        self.hsplit.setStretchFactor(1, 1)
        self.hsplit.setSizes([400, 820])
        self.setCentralWidget(self.hsplit)
        self.statusBar().showMessage("Ready")

    def finish_side(self, *extra_buttons):
        self.side.addWidget(self.run_btn)
        for b in extra_buttons:
            self.side.addWidget(b)
        ref = QPushButton("Refresh preview  (F5)")
        ref.clicked.connect(self.safe_refresh)
        self.side.addWidget(ref)
        self.side.addStretch(1)
        fit_panel(self.side.parentWidget())

    # ---- settings ------------------------------------------------------------
    def restore_frame(self, size=(1180, 760)):
        g = self.settings.value("geometry")
        if g is not None:
            self.restoreGeometry(g)
        else:
            self.resize(*size)
        for sp, key in ((self.hsplit, "hsplit"), (self.vsplit, "vsplit")):
            v = self.settings.value(key)
            if v is not None:
                sp.restoreState(v)

    def closeEvent(self, e):
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("hsplit", self.hsplit.saveState())
        self.settings.setValue("vsplit", self.vsplit.saveState())
        self.save_settings()
        self.disconnect_project()
        super().closeEvent(e)

    def save_settings(self):
        pass

    def disconnect_project(self):
        pass

    # ---- preview -------------------------------------------------------------
    def schedule(self, *args):
        if not self.running:
            self.timer.start()

    def safe_refresh(self):
        try:
            self.refresh()
        except Exception:
            self.say("Preview failed:\n" + traceback.format_exc())

    def refresh(self):
        pass

    # ---- log / run -------------------------------------------------------------
    def say(self, msg):
        self.log.appendPlainText(msg)
        self.log.ensureCursorVisible()

    def _live(self, line):
        self.say(line)
        QApplication.processEvents(QEventLoop.ExcludeUserInputEvents)  # show progress, no clicks

    def run_logged(self, label, fn, *args, **kwargs):
        """Call fn with its prints in the log; exceptions are logged, not raised."""
        self.running = True
        self.run_btn.setEnabled(False)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.say(f"\n===== {label} =====")
        self.statusBar().showMessage(label + "…")
        stream = LogStream(self._live)
        result = None
        try:
            with contextlib.redirect_stdout(stream):
                result = fn(*args, **kwargs)
            self.statusBar().showMessage(label + " - done", 8000)
        except Exception:
            self.say(traceback.format_exc())
            self.statusBar().showMessage(label + " - failed, see the log")
        finally:
            stream.flush()
            QApplication.restoreOverrideCursor()
            self.run_btn.setEnabled(True)
            self.running = False
        return result

    def run(self):
        pass
