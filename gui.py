"""
Fiber Trace Desktop GUI

What this does
- Lets users enter manual fields and pick two input files (JSON + CSV)
- Runs your existing pipeline (tries to import and call functions from your project)
- Saves outputs (KML + CSVs) to a chosen folder
- Previews the route on an embedded map when possible

Why PySide6?
- Native desktop UI that can be packaged to a single Windows .exe with PyInstaller
- Embedded Chromium (QtWebEngine) so we can show a Folium/Leaflet map preview

---
Project assumptions (adjust if your app.py differs)
- You already have modules like: _to_kml.py, kml_helper.py, parse_device_sheet.py,
  trace_report.py, fibre_trace.py, remove_add_algo.py and/or app.py orchestrating them.
- This GUI will try several call paths to run your core pipeline.
  If none are found, wire up `run_pipeline_core()` to your real functions (marked TODO).

Test with your sample files:
- You mentioned uploading samples such as /mnt/data/WO24218.json and /mnt/data/2.csv.
  The GUI includes a "Load sample paths" button that fills those in automatically if present.

Packaging to EXE (Windows):
- Create a venv, install deps from the requirements block below, then run the PyInstaller command.
- See the instructions at the bottom of this file.

Requirements (put these in requirements.txt):
    PySide6>=6.6
    PySide6-Addons>=6.6  # sometimes needed by QtWebEngine on some setups
    PySide6-Essentials>=6.6
    folium
    jinja2
    fastkml  # optional; if missing, preview falls back or is skipped
    shapely  # optional; fastkml may use it

"""
from __future__ import annotations
import os
import sys
import json
import traceback
from datetime import datetime
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QMessageBox,
    QLineEdit, QLabel, QPushButton, QHBoxLayout, QVBoxLayout, QFormLayout,
    QGroupBox, QSplitter, QCheckBox
)
from PySide6.QtCore import Qt, QUrl

# WebView for map preview (QtWebEngine)
try:
    from PySide6.QtWebEngineWidgets import QWebEngineView
except Exception:
    QWebEngineView = None  # Preview will be disabled if QtWebEngine isn't available

# Optional import of your existing pipeline modules (best-effort)
# These may or may not exist; failures are handled later.
try:
    import app as app_module  # your original orchestrator
except Exception:
    app_module = None

# Other helpers (import if present; it's fine if they are missing)
for modname in [
    "_to_kml", "kml_helper", "parse_device_sheet", "trace_report",
    "fibre_trace", "remove_add_algo"
]:
    try:
        __import__(modname)
    except Exception:
        pass

# ------------------------
# Pipeline adapter
# ------------------------

def run_pipeline_core(manual: dict, json_path: Path, csv_path: Path, out_dir: Path) -> dict:
    """Calls your existing processing code to produce outputs.

    This function tries a few likely call patterns:
      1) app_module.main(json_path, csv_path, out_dir, **manual)
      2) app_module.run(json_path, csv_path, out_dir, meta=manual)
      3) app_module.process(json_path, csv_path, out_dir, **manual)

    If none are available, REPLACE the TODO section with direct calls
    into your real functions (e.g., parse_device_sheet -> fibre_trace -> _to_kml -> trace_report).

    Returns a dict with keys like:
      {
        "kml": Path | None,
        "csvs": [Path, ...],
        "log": str
      }
    """
    log_lines = []
    produced_kml: Path | None = None
    produced_csvs: list[Path] = []

    def log(msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        log_lines.append(f"[{ts}] {msg}")

    json_path = Path(json_path)
    csv_path = Path(csv_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Try well-known orchestrators in your app.py
    if app_module is not None:
        for fname, kwargs in [
            ("main", dict(json_path=json_path, csv_path=csv_path, out_dir=out_dir, **manual)),
            ("run", dict(json_path=json_path, csv_path=csv_path, out_dir=out_dir, meta=manual)),
            ("process", dict(json_path=json_path, csv_path=csv_path, out_dir=out_dir, **manual)),
        ]:
            fn = getattr(app_module, fname, None)
            if callable(fn):
                log(f"Calling app.{fname}()")
                try:
                    result = fn(**kwargs)
                    # Result contract best-effort. If your function returns paths, normalize them.
                    if isinstance(result, dict):
                        produced_kml = Path(result.get("kml")) if result.get("kml") else None
                        csvs = result.get("csvs") or []
                        produced_csvs = [Path(p) for p in csvs if p]
                    # If result is a single path or tuple
                    elif isinstance(result, (str, os.PathLike)):
                        p = Path(result)
                        if p.suffix.lower() == ".kml":
                            produced_kml = p
                    elif isinstance(result, tuple):
                        for item in result:
                            if isinstance(item, (str, os.PathLike)) and str(item).lower().endswith('.kml'):
                                produced_kml = Path(item)
                            elif isinstance(item, (str, os.PathLike)) and str(item).lower().endswith('.csv'):
                                produced_csvs.append(Path(item))
                    log("Pipeline finished through app.py entrypoint.")
                    return {"kml": produced_kml, "csvs": produced_csvs, "log": "\n".join(log_lines)}
                except Exception:
                    log("app entrypoint raised an exception; falling back...")
                    log(traceback.format_exc())

    # TODO: If your project doesn't expose app.main/run/process, wire the exact calls below.
    # Example skeleton (replace with your real function names and signatures):
    try:
        log("Running fallback pipeline skeleton (adjust to your project).")
        # Example: parse CSV devices, compute trace, emit KML + reports
        # from parse_device_sheet import parse_devices
        # from fibre_trace import build_trace
        # from _to_kml import write_kml
        # from trace_report import write_reports
        # devices = parse_devices(csv_path)
        # trace = build_trace(json_path, devices, meta=manual)
        # produced_kml = out_dir / "trace.kml"
        # write_kml(trace, produced_kml)
        # produced_csvs = write_reports(trace, out_dir)
        # For now, we just copy inputs to outputs to keep the GUI usable until wired.
        import shutil
        produced_kml = out_dir / (json_path.stem + ".kml")
        shutil.copyfile(json_path, out_dir / json_path.name)
        shutil.copyfile(csv_path, out_dir / csv_path.name)
        with open(produced_kml, "w", encoding="utf-8") as f:
            f.write("<?xml version='1.0' encoding='UTF-8'?>\n<kml xmlns='http://www.opengis.net/kml/2.2'>\n<Document>\n<Name>Placeholder</Name>\n</Document>\n</kml>\n")
        produced_csvs = [out_dir / csv_path.name]
        log("Fallback pipeline wrote placeholder outputs. Replace with real functions.")
    except Exception:
        log("Fallback pipeline also failed:")
        log(traceback.format_exc())

    return {"kml": produced_kml, "csvs": produced_csvs, "log": "\n".join(log_lines)}


# ------------------------
# Map preview helpers
# ------------------------

def kml_to_map_html(kml_path: Path, html_out: Path) -> Path | None:
    """Renders a very simple Folium map and tries to overlay the KML.
    If fastkml is not available or KML has complex geometries, we still
    give the user a base map centered roughly on Toronto.
    """
    try:
        import folium
        from folium.plugins import BeautifyIcon  # noqa: F401  (not required but keeps plugin import tested)
    except Exception:
        return None

    # Create a basic map; if we can parse the KML, we will fit to bounds
    m = folium.Map(location=[43.65, -79.38], zoom_start=11)

    # Try to add KML overlay. Folium doesn't parse KML natively; we convert to GeoJSON if possible.
    try:
        from fastkml import kml
        from shapely.geometry import mapping
        k = kml.KML()
        with open(kml_path, 'rb') as f:
            k.from_string(f.read())
        geoms = []
        def recurse(feat):
            for ftr in getattr(feat, 'features', []):
                if hasattr(ftr, 'geometry') and ftr.geometry is not None:
                    geoms.append(ftr.geometry)
                recurse(ftr)
        recurse(k)
        gj = {
            "type": "FeatureCollection",
            "features": [
                {"type": "Feature", "geometry": mapping(g), "properties": {}} for g in geoms
            ],
        }
        folium.GeoJson(gj, name="Trace").add_to(m)
        try:
            # fit to bounds if we have any
            bounds = []
            for g in geoms:
                try:
                    minx, miny, maxx, maxy = g.bounds
                    bounds.append([[miny, minx], [maxy, maxx]])
                except Exception:
                    pass
            if bounds:
                # Compute overall bounds
                min_lat = min(b[0][0] for b in bounds)
                min_lon = min(b[0][1] for b in bounds)
                max_lat = max(b[1][0] for b in bounds)
                max_lon = max(b[1][1] for b in bounds)
                m.fit_bounds([[min_lat, min_lon], [max_lat, max_lon]])
        except Exception:
            pass
    except Exception:
        # No overlay, just base map
        pass

    m.save(str(html_out))
    return html_out if html_out.exists() else None


# ------------------------
# GUI
# ------------------------
class FiberTraceWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Fiber Trace – Desktop GUI")
        self.resize(1200, 800)

        self.out_dir: Path | None = None
        self.html_preview_path: Path | None = None

        root = QWidget()
        self.setCentralWidget(root)

        # Left side: form + actions
        form_group = QGroupBox("Inputs")
        form = QFormLayout()

        self.project_edit = QLineEdit()
        self.work_order_edit = QLineEdit()
        self.tech_edit = QLineEdit()
        self.date_edit = QLineEdit(datetime.now().strftime("%Y-%m-%d"))

        self.json_edit = QLineEdit()
        self.csv_edit = QLineEdit()
        self.json_btn = QPushButton("Browse JSON…")
        self.csv_btn = QPushButton("Browse CSV…")
        self.json_btn.clicked.connect(self.pick_json)
        self.csv_btn.clicked.connect(self.pick_csv)

        self.sample_btn = QPushButton("Load sample paths")
        self.sample_btn.clicked.connect(self.load_samples)

        self.out_btn = QPushButton("Choose output folder…")
        self.out_btn.clicked.connect(self.pick_out_dir)
        self.out_label = QLabel("<i>No output folder chosen</i>")

        form.addRow("Project:", self.project_edit)
        form.addRow("Work Order:", self.work_order_edit)
        form.addRow("Technician:", self.tech_edit)
        form.addRow("Date (YYYY-MM-DD):", self.date_edit)

        jrow = QHBoxLayout(); jrow.addWidget(self.json_edit); jrow.addWidget(self.json_btn)
        crow = QHBoxLayout(); crow.addWidget(self.csv_edit); crow.addWidget(self.csv_btn)
        form.addRow("Input JSON:", wrap_layout(jrow))
        form.addRow("Input CSV:", wrap_layout(crow))
        form.addRow(self.sample_btn)
        form.addRow(self.out_btn)
        form.addRow(self.out_label)
        form_group.setLayout(form)

        run_group = QGroupBox("Actions")
        v = QVBoxLayout()
        self.preview_check = QCheckBox("Generate map preview (if possible)")
        self.preview_check.setChecked(True)
        self.run_btn = QPushButton("Run & Save Outputs")
        self.run_btn.clicked.connect(self.run_pipeline)
        v.addWidget(self.preview_check)
        v.addWidget(self.run_btn)
        v.addStretch(1)
        run_group.setLayout(v)

        left_panel = QVBoxLayout()
        left_panel.addWidget(form_group)
        left_panel.addWidget(run_group)
        left_panel.addStretch(1)
        left = QWidget(); left.setLayout(left_panel)

        # Right side: log + preview
        right_box = QVBoxLayout()

        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("Logs will appear here…")
        right_box.addWidget(QLabel("Log"))
        right_box.addWidget(self.log, 2)

        if QWebEngineView is not None:
            self.web = QWebEngineView()
            right_box.addWidget(QLabel("Preview"))
            right_box.addWidget(self.web, 5)
        else:
            self.web = None
            msg = QLabel("QtWebEngine not available — map preview disabled.")
            msg.setStyleSheet("color:#a33")
            right_box.addWidget(msg)

        right = QWidget(); right.setLayout(right_box)

        splitter = QSplitter(); splitter.addWidget(left); splitter.addWidget(right); splitter.setStretchFactor(1, 1)
        layout = QVBoxLayout(); layout.addWidget(splitter)
        root.setLayout(layout)

    # ---------- UI helpers ----------
    def append_log(self, text: str):
        self.log.appendPlainText(text)
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    def pick_json(self):
        fn, _ = QFileDialog.getOpenFileName(self, "Select JSON", str(Path.cwd()), "JSON Files (*.json)")
        if fn:
            self.json_edit.setText(fn)

    def pick_csv(self):
        fn, _ = QFileDialog.getOpenFileName(self, "Select CSV", str(Path.cwd()), "CSV Files (*.csv)")
        if fn:
            self.csv_edit.setText(fn)

    def pick_out_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Choose output folder", str(Path.cwd()))
        if d:
            self.out_dir = Path(d)
            self.out_label.setText(str(self.out_dir))

    def load_samples(self):
        # Fills in the known sample paths if they exist
        candidates = [
            Path("/mnt/data/WO24218.json"), Path("WO24218.json"),
            Path.cwd() / "WO24218.json"
        ]
        for c in candidates:
            if c.exists():
                self.json_edit.setText(str(c)); break
        csv_candidates = [
            Path("/mnt/data/2.csv"), Path("/mnt/data/WO24218.csv"), Path("2.csv"),
            Path.cwd() / "2.csv"
        ]
        for c in csv_candidates:
            if c.exists():
                self.csv_edit.setText(str(c)); break
        if self.out_dir is None:
            default = Path.cwd() / "outputs"
            self.out_dir = default
            self.out_label.setText(str(default))

    # ---------- Run ----------
    def run_pipeline(self):
        try:
            json_path = Path(self.json_edit.text().strip())
            csv_path = Path(self.csv_edit.text().strip())
            if not json_path.exists():
                raise FileNotFoundError(f"JSON not found: {json_path}")
            if not csv_path.exists():
                raise FileNotFoundError(f"CSV not found: {csv_path}")

            if self.out_dir is None:
                self.out_dir = Path.cwd() / "outputs"
                self.out_label.setText(str(self.out_dir))
            self.out_dir.mkdir(parents=True, exist_ok=True)

            manual = {
                "project": self.project_edit.text().strip(),
                "work_order": self.work_order_edit.text().strip(),
                "technician": self.tech_edit.text().strip(),
                "date": self.date_edit.text().strip(),
            }

            self.append_log("Starting pipeline…")
            res = run_pipeline_core(manual, json_path, csv_path, self.out_dir)
            self.append_log(res.get("log", "(no log)"))

            kml_path = res.get("kml")
            csvs = res.get("csvs") or []
            if kml_path:
                self.append_log(f"KML: {kml_path}")
            for c in csvs:
                self.append_log(f"CSV:  {c}")

            # Preview
            if self.preview_check.isChecked() and kml_path and QWebEngineView is not None:
                html_out = self.out_dir / "preview.html"
                html = kml_to_map_html(kml_path, html_out)
                if html:
                    self.html_preview_path = html
                    self.web.setUrl(QUrl.fromLocalFile(str(html)))
                    self.append_log("Preview updated.")
                else:
                    self.append_log("Preview skipped (folium/fastkml not available).")
            elif not kml_path:
                self.append_log("No KML produced — preview skipped.")

            QMessageBox.information(self, "Done", "Processing finished. Outputs are in:\n" + str(self.out_dir))
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(self, "Error", f"{e}\n\nSee log for details.")
            self.append_log("ERROR:\n" + traceback.format_exc())


def wrap_layout(layout: QtWidgets.QLayout) -> QWidget:
    w = QWidget(); w.setLayout(layout); return w


def main():
    app = QApplication(sys.argv)
    win = FiberTraceWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

"""
PACKAGING NOTES (Windows .exe)
==============================
1) Create virtual environment & install deps
-------------------------------------------
python -m venv .venv
.venv\\Scripts\\activate
pip install -U pip
pip install -r requirements.txt

2) PyInstaller command (single file exe)
---------------------------------------
# Run from the folder containing this file and your modules
pyinstaller \
  --noconfirm --clean \
  --name "FiberTrace" \
  --onefile \
  --add-data "templates;templates" \
  --add-data "*.py;." \
  --hidden-import PySide6.QtWebEngineWidgets \
  gui_fibertrace.py

Notes:
- Replace gui_fibertrace.py with the real filename of this script.
- If your pipeline needs extra data folders (icons, html, etc.), add them with --add-data.
- On some setups, QtWebEngine needs extra resources. If preview fails in the .exe,
  switch to --onefolder and let PyInstaller collect QtWebEngineProcess and resources automatically:

pyinstaller \
  --noconfirm --clean \
  --name "FiberTrace" \
  --onefolder \
  --hidden-import PySide6.QtWebEngineWidgets \
  gui_fibertrace.py

3) Ship
-------
Distribute the dist/FiberTrace.exe (or the whole dist/FiberTrace folder when using --onefolder).

Troubleshooting
---------------
- If app.py exposes a different entrypoint/signature, edit run_pipeline_core() accordingly.
- If Folium or fastkml are not desired, uncheck preview in the UI or remove those deps.
- If the web preview crashes inside the exe, use --onefolder first to confirm QtWebEngine resources are included.
"""
