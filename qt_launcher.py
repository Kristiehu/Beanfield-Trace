# qt_launcher.py
import os, sys, threading, time, socket
from pathlib import Path

APP_TITLE = "Beanfield Trace App"
PORT = 8501
WIDTH, HEIGHT = 1200, 800

BASE = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
APP_PY = os.path.join(BASE, "app", "app.py")

def port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", port)) == 0

def run_streamlit():
    # run Streamlit in-process
    try:
        from streamlit.web import cli as stcli
    except Exception:
        import streamlit.web.cli as stcli
    sys.argv = [
        "streamlit", "run", APP_PY,
        "--server.port", str(PORT),
        "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
        "--client.showErrorDetails", "false",
    ]
    stcli.main()

# Start Streamlit server in a background thread
if not port_in_use(PORT):
    t = threading.Thread(target=run_streamlit, daemon=True)
    t.start()

# Optional: QtWebEngine sandbox can be finicky in frozen apps—disable sandbox
os.environ.setdefault("QTWEBENGINE_DISABLE_SANDBOX", "1")

# ---- Qt window embedding Chromium ----
from PyQt6.QtCore import QUrl, Qt
from PyQt6.QtWidgets import QApplication
from PyQt6.QtWebEngineWidgets import QWebEngineView

app = QApplication(sys.argv)

# Wait for server to come up
for _ in range(120):
    if port_in_use(PORT):
        break
    time.sleep(0.25)

view = QWebEngineView()
view.setWindowTitle(APP_TITLE)
view.resize(WIDTH, HEIGHT)
view.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
view.setUrl(QUrl(f"http://127.0.0.1:{PORT}"))
view.show()

sys.exit(app.exec())
