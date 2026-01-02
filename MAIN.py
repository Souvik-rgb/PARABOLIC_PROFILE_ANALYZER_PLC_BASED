import sys, os, json
import pandas as pd
import serial.tools.list_ports

from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QHBoxLayout,
    QLineEdit, QTableWidget, QTableWidgetItem, QStackedWidget, QMenuBar,
    QFileDialog, QInputDialog, QDialog, QComboBox, QMessageBox
)
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QPixmap

from pymodbus.client.sync import ModbusSerialClient
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
from matplotlib.figure import Figure


# ================= THEME =================
BRAND_COLOR = "#2A82DA"

DARK_THEME = f"""
QWidget {{ background:#121417; color:#E6E6E6; font-family:Segoe UI; }}
QMenuBar {{ background:#1B1F24; }}
QMenuBar::item:selected {{ background:{BRAND_COLOR}; }}
QMenu {{ background:#1B1F24; }}
QPushButton {{ background:{BRAND_COLOR}; color:white; padding:8px 14px; border-radius:6px; }}
QLineEdit,QTableWidget {{ background:#1B1F24; border:1px solid #333; }}
QHeaderView::section {{ background:#242A31; font-weight:bold; }}
"""

LIGHT_THEME = f"""
QWidget {{ background:#F4F6F8; color:#111; font-family:Segoe UI; }}
QMenuBar {{ background:white; }}
QMenuBar::item:selected {{ background:{BRAND_COLOR}; color:white; }}
QMenu {{ background:white; }}
QPushButton {{ background:{BRAND_COLOR}; color:white; padding:8px 14px; border-radius:6px; }}
QLineEdit,QTableWidget {{ background:white; border:1px solid #CCC; }}
QHeaderView::section {{ background:#EAEAEA; font-weight:bold; }}
"""


# ================= CONFIG =================
CONFIG = {
    "baudrate": 9600,

    "x_encoder": {"type": "D", "addr": 100},
    "y_encoder": {"type": "D", "addr": 102},

    "start_bit": {"type": "M", "addr": 10},
    "stop_bit": {"type": "M", "addr": 11},
    "run_bit": {"type": "M", "addr": 12},
    "plot_start_bit": {"type": "M", "addr": 13},

    "x_zero_bit": {"type": "M", "addr": 30},
    "y_zero_bit": {"type": "M", "addr": 31},

    "x_ppr": 1000,
    "y_ppr": 1000
}

PASSWORD = "1234"
client = None
connected = False
current_project = None


# ================= PLC HELPERS =================
def auto_port():
    for p in serial.tools.list_ports.comports():
        if "USB" in p.description or "Serial" in p.description:
            return p.device
    return None


def connect_plc():
    global client, connected
    port = auto_port()
    if not port:
        connected = False
        return

    client = ModbusSerialClient(
        method="rtu",
        port=port,
        baudrate=CONFIG["baudrate"],
        parity="E",
        stopbits=1,
        bytesize=8,
        timeout=1
    )
    connected = client.connect()


def read_plc(reg):
    if not connected:
        return None

    t, a = reg["type"], reg["addr"]

    if t in ["D", "C", "T"]:
        r = client.read_holding_registers(a, 1, unit=1)
        return r.registers[0] if not r.isError() else None

    if t in ["M", "Y"]:
        r = client.read_coils(a, 1, unit=1)
        return r.bits[0] if not r.isError() else None

    if t == "X":
        r = client.read_discrete_inputs(a, 1, unit=1)
        return r.bits[0] if not r.isError() else None

    return None


def write_plc_bit(reg):
    if connected and reg["type"] in ["M", "Y"]:
        client.write_coil(reg["addr"], True, unit=1)


def check_password():
    p, ok = QInputDialog.getText(None, "Password", "Enter Password", QLineEdit.Password)
    return ok and p == PASSWORD


def restart_app():
    os.execv(sys.executable, [sys.executable] + sys.argv)


# ================= CREATE PROFILE =================
class CreateProfileDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Create Reference Profile")
        self.resize(520, 420)

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Reference X", "Reference Y"])

        add = QPushButton("Add Row")
        save = QPushButton("Save CSV")
        use = QPushButton("Use Profile")

        add.clicked.connect(lambda: self.table.insertRow(self.table.rowCount()))
        save.clicked.connect(self.save_csv)
        use.clicked.connect(self.accept)

        btns = QHBoxLayout()
        btns.addWidget(add)
        btns.addStretch()
        btns.addWidget(save)
        btns.addWidget(use)

        layout = QVBoxLayout(self)
        layout.addWidget(self.table)
        layout.addLayout(btns)

    def get_profile(self):
        rx, ry = [], []
        for i in range(self.table.rowCount()):
            x = self.table.item(i, 0)
            y = self.table.item(i, 1)
            if x and y:
                try:
                    rx.append(float(x.text()))
                    ry.append(float(y.text()))
                except ValueError:
                    pass
        return rx, ry

    def save_csv(self):
        rx, ry = self.get_profile()
        if not rx:
            QMessageBox.warning(self, "Empty", "No data to save")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save CSV", "", "CSV (*.csv)")
        if path:
            pd.DataFrame({"Reference_X": rx, "Reference_Y": ry}).to_csv(path, index=False)
            QMessageBox.information(self, "Saved", "Profile saved successfully")


# ================= MAIN WINDOW =================
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Parabolic Leaf Profile Analyzer")
        self.resize(1400, 850)
        self.dark = True

        connect_plc()

        self.stack = QStackedWidget()
        self.live = LivePage()
        self.settings = SettingsPage(self)
        self.more = MoreSettingsPage(self)

        self.stack.addWidget(self.live)
        self.stack.addWidget(self.settings)
        self.stack.addWidget(self.more)

        self.menu = QMenuBar()
        self.build_menu()

        layout = QVBoxLayout(self)
        layout.setMenuBar(self.menu)
        layout.addWidget(self.stack)

    def build_menu(self):
        file = self.menu.addMenu("File")
        file.addAction("New").triggered.connect(restart_app)
        file.addAction("Open").triggered.connect(self.open_file)
        file.addAction("Save").triggered.connect(self.save_project)
        file.addAction("Save As").triggered.connect(self.saveas_project)
        file.addSeparator()
        file.addAction("Exit").triggered.connect(self.close)

        settings = self.menu.addMenu("Settings")
        settings.addAction("Settings").triggered.connect(self.goto_settings)
        settings.addAction("More Settings").triggered.connect(self.goto_more)

        profile = self.menu.addMenu("Create Profile")
        profile.addAction("Create Table").triggered.connect(self.create_profile)

        view = self.menu.addMenu("View")
        view.addAction("Toggle Light / Dark").triggered.connect(self.toggle_theme)

    def toggle_theme(self):
        self.dark = not self.dark
        QApplication.instance().setStyleSheet(DARK_THEME if self.dark else LIGHT_THEME)

    def create_profile(self):
        dlg = CreateProfileDialog(self)
        if dlg.exec_():
            rx, ry = dlg.get_profile()
            self.live.set_reference(rx, ry)

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open File", "", "CSV (*.csv);;Project (*.json)"
        )
        if path.endswith(".csv"):
            df = pd.read_csv(path)
            self.live.set_reference(df.iloc[:, 0].tolist(), df.iloc[:, 1].tolist())
        elif path.endswith(".json"):
            with open(path, "r") as f:
                CONFIG.update(json.load(f))

    def save_project(self):
        global current_project
        if current_project:
            with open(current_project, "w") as f:
                json.dump(CONFIG, f, indent=4)
        else:
            self.saveas_project()

    def saveas_project(self):
        global current_project
        path, _ = QFileDialog.getSaveFileName(self, "Save Project", "", "Project (*.json)")
        if path:
            current_project = path
            with open(path, "w") as f:
                json.dump(CONFIG, f, indent=4)

    def goto_settings(self):
        if check_password():
            self.stack.setCurrentIndex(1)

    def goto_more(self):
        if check_password():
            self.stack.setCurrentIndex(2)


# ================= LIVE PAGE =================
class LivePage(QWidget):
    def __init__(self):
        super().__init__()

        self.ref_x, self.ref_y = [], []
        self.act_x, self.act_y = [], []

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.read_plc)

        self.logo = QLabel()
        self.logo.setPixmap(QPixmap(resource_path("logo.png")).scaledToHeight(36))


        self.run = QLabel("● RUN OFF")
        self.run.setStyleSheet("color:gray;font-weight:bold")

        self.blink = False
        self.blink_timer = QTimer(self)
        self.blink_timer.timeout.connect(self.animate_run)
        self.blink_timer.start(500)

        self.xv = QLabel("X: -")
        self.yv = QLabel("Y: -")

        start = QPushButton("START")
        stop = QPushButton("STOP")
        start.clicked.connect(lambda: self.timer.start(100))
        stop.clicked.connect(self.timer.stop)

        header = QHBoxLayout()
        header.addWidget(self.logo.)
        header.addWidget(start)
        header.addWidget(stop)
        header.addWidget(self.run)
        header.addStretch()
        header.addWidget(self.xv)
        header.addWidget(self.yv)

        self.fig = Figure()
        self.canvas = FigureCanvasQTAgg(self.fig)
        self.ax = self.fig.add_subplot(111)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["Ref X", "Ref Y", "Actual X", "Actual Y", "Diff X", "Diff Y"]
        )

        layout = QVBoxLayout(self)
        layout.addLayout(header)
        layout.addWidget(self.canvas)
        layout.addWidget(self.table)

    def set_reference(self, rx, ry):
        self.ref_x = list(rx)
        self.ref_y = list(ry)
        self.update_plot()
        self.update_table()

    def animate_run(self):
        if "ON" in self.run.text():
            self.blink = not self.blink
            self.run.setStyleSheet(
                f"color:{'#00FF6A' if self.blink else '#008F3A'};font-weight:bold"
            )

    def read_plc(self):
        xr = read_plc(CONFIG["x_encoder"])
        yr = read_plc(CONFIG["y_encoder"])
        run = read_plc(CONFIG["run_bit"])

        if xr is None or yr is None:
            return

        ax = xr / CONFIG["x_ppr"]
        ay = yr / CONFIG["y_ppr"]

        self.act_x.append(ax)
        self.act_y.append(ay)

        self.xv.setText(f"X: {ax:.3f}")
        self.yv.setText(f"Y: {ay:.3f}")
        self.run.setText("● RUN ON" if run else "● RUN OFF")

        self.update_plot()
        self.update_table()

    def update_plot(self):
        self.ax.clear()
        if self.ref_x:
            self.ax.plot(self.ref_x, self.ref_y, "g--", linewidth=2, label="Reference")
        if self.act_x:
            self.ax.plot(self.act_x, self.act_y, "b", linewidth=2, label="Actual")
        self.ax.legend()
        self.ax.grid(True)
        self.canvas.draw_idle()

    def update_table(self):
        rows = max(len(self.ref_x), len(self.act_x))
        self.table.setRowCount(rows)

        for i in range(rows):
            rx = self.ref_x[i] if i < len(self.ref_x) else ""
            ry = self.ref_y[i] if i < len(self.ref_y) else ""
            ax = self.act_x[i] if i < len(self.act_x) else ""
            ay = self.act_y[i] if i < len(self.act_y) else ""

            dx = ax - rx if ax != "" and rx != "" else ""
            dy = ay - ry if ay != "" and ry != "" else ""

            self.table.setItem(i, 0, QTableWidgetItem(str(rx)))
            self.table.setItem(i, 1, QTableWidgetItem(str(ry)))
            self.table.setItem(i, 2, QTableWidgetItem(str(ax)))
            self.table.setItem(i, 3, QTableWidgetItem(str(ay)))
            self.table.setItem(i, 4, QTableWidgetItem(str(dx)))
            self.table.setItem(i, 5, QTableWidgetItem(str(dy)))


# ================= SETTINGS PAGE =================
class SettingsPage(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.fields = {}

        layout = QVBoxLayout(self)

        def row(title, key):
            h = QHBoxLayout()
            t = QComboBox()
            t.addItems(["D", "M", "X", "Y", "C", "T"])
            t.setCurrentText(CONFIG[key]["type"])
            n = QLineEdit(str(CONFIG[key]["addr"]))
            self.fields[key] = (t, n)
            h.addWidget(QLabel(title))
            h.addWidget(t)
            h.addWidget(n)
            layout.addLayout(h)

        row("X Encoder", "x_encoder")
        row("Y Encoder", "y_encoder")
        row("START Bit", "start_bit")
        row("STOP Bit", "stop_bit")
        row("RUN Bit", "run_bit")
        row("PLOT START Bit", "plot_start_bit")

        save = QPushButton("SAVE")
        save.clicked.connect(self.save)
        layout.addWidget(save)

    def save(self):
        for k, (t, n) in self.fields.items():
            CONFIG[k] = {"type": t.currentText(), "addr": int(n.text())}
        self.parent.stack.setCurrentIndex(0)


# ================= MORE SETTINGS =================
class MoreSettingsPage(QWidget):
    def __init__(self, parent):
        super().__init__()
        layout = QVBoxLayout(self)

        self.xp = QLineEdit(str(CONFIG["x_ppr"]))
        self.yp = QLineEdit(str(CONFIG["y_ppr"]))

        zx = QPushButton("ZERO X")
        zy = QPushButton("ZERO Y")

        zx.clicked.connect(lambda: write_plc_bit(CONFIG["x_zero_bit"]))
        zy.clicked.connect(lambda: write_plc_bit(CONFIG["y_zero_bit"]))

        layout.addWidget(QLabel("X PPR"))
        layout.addWidget(self.xp)
        layout.addWidget(QLabel("Y PPR"))
        layout.addWidget(self.yp)
        layout.addWidget(zx)
        layout.addWidget(zy)

        save = QPushButton("SAVE")
        save.clicked.connect(self.save)
        layout.addWidget(save)

    def save(self):
        CONFIG["x_ppr"] = int(self.xp.text())
        CONFIG["y_ppr"] = int(self.yp.text())


# ================= MAIN =================
app = QApplication(sys.argv)
app.setStyleSheet(DARK_THEME)
win = MainWindow()
win.show()
sys.exit(app.exec_())


