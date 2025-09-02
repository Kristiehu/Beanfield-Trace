import os, sys, subprocess, socket, time, signal, shutil, tempfile
from pathlib import Path

# ---------------- CONFIG ----------------
APP_REL_PATH = Path("app/app.py")
STREAMLIT_PORT = 8501
LOCK_PORT = 8765                 # single-instance lock port
WINDOW_TITLE = "My Streamlit App"
WINDOW_W, WINDOW_H = 1200, 800
STARTUP_TIMEOUT_S = 40
# ----------------------------------------

def _resource_path(rel: Path) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    return (base / rel).resolve()

def port_in_use(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.2)
        return s.connect_ex(("127.0.0.1", port)) == 0

def acquire_single_instance_lock(port: int):
    """Prevents multiple launcher instances (returns the bound socket)."""
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", port))
        s.listen(1)
        return s  # keep it open for the process lifetime
    except OSError:
        raise SystemExit("Another instance is already running.")

def launch_streamlit(app_path: Path, port: int) -> subprocess.Popen:
    env = os.environ.copy()
    env["BROWSER"] = "none"                    # prevent default browser
    env["STREAMLIT_BROWSER_GATHERUSAGESTATS"] = "false"
    # Optional: choose a webview backend explicitly on macOS to avoid extra processes
    if sys.platform == "darwin":
        env.setdefault("PYWEBVIEW_GUI", "cocoa")

    cmd = [
        sys.executable, "-m", "streamlit", "run", str(app_path),
        "--server.port", str(port),
        "--server.headless", "true",
        "--server.fileWatcherType", "none",      # <- prevent reload cascades
        "--logger.level", "error",
    ]
    return subprocess.Popen(cmd, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

def wait_for_server(port: int, timeout=STARTUP_TIMEOUT_S):
    t0 = time.time()
    while time.time() - t0 < timeout:
        if port_in_use(port):
            return True
        time.sleep(0.2)
    return False

def main():
    # Single-instance lock
    lock_socket = acquire_single_instance_lock(LOCK_PORT)

    # Resolve app path; when frozen, copy to a temp dir so Streamlit can write
    app_path = _resource_path(APP_REL_PATH)
    if getattr(sys, "_MEIPASS", None):
        tmpdir = Path(tempfile.mkdtemp(prefix="st_app_"))
        root = APP_REL_PATH.parts[0]  # "app"
        shutil.copytree(_resource_path(Path(root)), tmpdir / root, dirs_exist_ok=True)
        app_path = tmpdir / APP_REL_PATH

    # Don’t start a second Streamlit if it’s already up (e.g., previous crash left it running)
    if not port_in_use(STREAMLIT_PORT):
        proc = launch_streamlit(app_path, STREAMLIT_PORT)
    else:
        proc = None  # reuse existing server

    ok = wait_for_server(STREAMLIT_PORT)
    if not ok:
        if proc is not None:
            try:
                out = proc.stdout.read().decode(errors="ignore")
                print(out)
            except Exception:
                pass
        raise SystemExit("Streamlit server failed to start on time.")

    # Create a single native window and hand over control
    import webview
    window = webview.create_window(WINDOW_TITLE, f"http://127.0.0.1:{STREAMLIT_PORT}",
                                   width=WINDOW_W, height=WINDOW_H)
    try:
        # Start GUI loop; no debug, no http_server
        webview.start()
    finally:
        # Clean shutdown of the child server if we launched it
        if proc and proc.poll() is None:
            if sys.platform == "win32":
                proc.terminate()
            else:
                os.kill(proc.pid, signal.SIGTERM)
            try:
                proc.wait(timeout=5)
            except Exception:
                proc.kill()
        try:
            lock_socket.close()
        except Exception:
            pass

if __name__ == "__main__":
    # macOS / PyInstaller safety to avoid re-exec loops
    try:
        import multiprocessing as mp
        mp.freeze_support()
    except Exception:
        pass
    main()
