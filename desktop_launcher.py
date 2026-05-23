import os
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from pathlib import Path
from queue import Empty, Queue


LOCAL_HOST = "localhost"
LOCAL_CONNECT_HOST = "127.0.0.1"
PORT_SEARCH_ATTEMPTS = 10
CHILD_ENV = "EINSATZBERICHT_STREAMLIT_CHILD"
PROCESS_START_TIME = time.perf_counter()


class BootstrapSplash:
    def __init__(self) -> None:
        self._events: Queue[tuple[str, str | int]] = Queue()
        self._ready = threading.Event()
        self._thread: threading.Thread | None = None

    def start(self, message: str) -> None:
        if self._thread:
            return
        self._thread = threading.Thread(target=self._run, args=(message,), daemon=True)
        self._thread.start()
        self._ready.wait(timeout=1.5)

    def update(self, message: str) -> None:
        self._events.put(("update", message))

    def close(self, delay_ms: int = 0) -> None:
        self._events.put(("close", max(int(delay_ms), 0)))

    def _run(self, message: str) -> None:
        try:
            import tkinter as tk
            from tkinter import ttk

            root = tk.Tk()
            root.title("Einsatzbericht Manager")
            root.resizable(False, False)
            root.protocol("WM_DELETE_WINDOW", lambda: None)

            width = 460
            height = 170
            root.update_idletasks()
            x = max((root.winfo_screenwidth() - width) // 2, 0)
            y = max((root.winfo_screenheight() - height) // 2, 0)
            root.geometry(f"{width}x{height}+{x}+{y}")

            container = ttk.Frame(root, padding=24)
            container.pack(fill="both", expand=True)

            ttk.Label(
                container,
                text="Einsatzbericht Manager wird gestartet",
                font=("Segoe UI", 12, "bold"),
            ).pack(anchor="w")

            status_var = tk.StringVar(value=message)
            ttk.Label(container, textvariable=status_var, wraplength=400).pack(anchor="w", pady=(12, 16))

            progress = ttk.Progressbar(container, mode="indeterminate")
            progress.pack(fill="x")
            progress.start(12)

            def pump_events() -> None:
                try:
                    while True:
                        event, value = self._events.get_nowait()
                        if event == "update":
                            status_var.set(str(value))
                        elif event == "close":
                            root.after(int(value), root.destroy)
                            return
                except Empty:
                    pass
                root.after(100, pump_events)

            self._ready.set()
            root.after(100, pump_events)
            root.mainloop()
        except Exception:
            self._ready.set()


def _current_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _child_command() -> list[str]:
    if getattr(sys, "frozen", False):
        return [sys.executable, "--streamlit-child"]
    return [sys.executable, str(Path(__file__).resolve()), "--streamlit-child"]


def _port_accepts_connection(port: int, timeout: float = 0.01) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(timeout)
        return sock.connect_ex((LOCAL_CONNECT_HOST, port)) == 0


def _server_ready(port: int, timeout: float = 0.35) -> bool:
    if not _port_accepts_connection(port):
        return False
    try:
        with urllib.request.urlopen(f"http://{LOCAL_CONNECT_HOST}:{port}/_stcore/health", timeout=timeout):
            return True
    except Exception:
        return False


def _candidate_ports(preferred: int = 8501, attempts: int = PORT_SEARCH_ATTEMPTS) -> list[int]:
    return list(range(preferred, preferred + attempts))


def _startup_log_path(runtime: dict) -> Path:
    log_dir = runtime["user_root"] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir / "startup_timing.log"


def _log_startup_event(runtime: dict, label: str, start_time: float = PROCESS_START_TIME) -> None:
    try:
        elapsed = time.perf_counter() - start_time
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        role = "child" if os.environ.get(CHILD_ENV) == "1" or "--streamlit-child" in sys.argv else "bootstrap"
        _startup_log_path(runtime).open("a", encoding="utf-8").write(
            f"{timestamp}\tpid={os.getpid()}\trole={role}\t+{elapsed:.3f}s\t{label}\n"
        )
    except Exception:
        pass


def _open_existing_instance(runtime: dict) -> bool:
    for port in _candidate_ports():
        if _server_ready(int(port)):
            webbrowser.open(f"http://{LOCAL_HOST}:{int(port)}")
            return True
    return False


def _wait_for_child(runtime: dict, child: subprocess.Popen, splash: BootstrapSplash, timeout_seconds: int = 75) -> bool:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        splash.update("Browser wird geoeffnet, sobald die App bereit ist...")
        if _open_existing_instance(runtime):
            splash.update("App ist bereit. Browser wurde geoeffnet.")
            _log_startup_event(runtime, "child health check succeeded; browser opened")
            return True
        if child.poll() is not None:
            _log_startup_event(runtime, f"child exited before health check succeeded: code={child.returncode}")
            return False
        time.sleep(0.5)
    _log_startup_event(runtime, "child health check timed out")
    return False


def _start_child(runtime: dict) -> subprocess.Popen:
    env = os.environ.copy()
    env[CHILD_ENV] = "1"
    env["EINSATZBERICHT_BOOTSTRAP_PARENT"] = "1"

    log_dir = runtime["user_root"] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    stdout = (log_dir / "launch_stdout.log").open("ab")
    stderr = (log_dir / "launch_stderr.log").open("ab")
    kwargs = {
        "cwd": str(_current_dir()),
        "env": env,
        "stdout": stdout,
        "stderr": stderr,
        "close_fds": False,
    }
    if sys.platform == "win32":
        kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return subprocess.Popen(_child_command(), **kwargs)


def run_bootstrap() -> int:
    os.chdir(_current_dir())
    splash = BootstrapSplash()
    splash.start("App-Prozess wird vorbereitet...")

    try:
        from desktop_runtime import (
            configure_streamlit_runtime,
            maybe_check_for_updates,
            prepare_runtime_environment,
            report_startup_failure,
            show_pending_update_changelog,
        )

        configure_streamlit_runtime()
        runtime = prepare_runtime_environment()
        _log_startup_event(runtime, "bootstrap runtime prepared")

        splash.update("Updates werden geprueft...")
        update_start = time.perf_counter()
        if maybe_check_for_updates():
            _log_startup_event(runtime, f"update check started installer after {time.perf_counter() - update_start:.3f}s")
            splash.close()
            return 0
        show_pending_update_changelog()
        _log_startup_event(runtime, f"update check finished after {time.perf_counter() - update_start:.3f}s")

        splash.update("Laufende Instanz wird gesucht...")
        existing_start = time.perf_counter()
        if _open_existing_instance(runtime):
            _log_startup_event(runtime, f"existing instance found after {time.perf_counter() - existing_start:.3f}s")
            splash.close(delay_ms=500)
            time.sleep(0.6)
            return 0
        _log_startup_event(runtime, f"existing instance search finished after {time.perf_counter() - existing_start:.3f}s")

        splash.update("Streamlit wird gestartet...")
        child = _start_child(runtime)
        _log_startup_event(runtime, f"child process started: pid={child.pid}")
        if _wait_for_child(runtime, child, splash):
            splash.close(delay_ms=1500)
            time.sleep(1.6)
            return 0

        splash.close()
        raise RuntimeError("Streamlit wurde gestartet, ist aber nicht erreichbar oder wurde beendet.")
    except Exception as exc:
        splash.close()
        try:
            from desktop_runtime import report_startup_failure

            report_startup_failure(exc)
        except Exception:
            pass
        return 1


def _apply_streamlit_options() -> None:
    from streamlit import config as st_config

    for option_name, value in {
        "global.developmentMode": False,
        "server.headless": True,
        "server.showEmailPrompt": False,
        "browser.gatherUsageStats": False,
    }.items():
        try:
            st_config.set_option(option_name, value)
        except Exception:
            pass


def _find_available_port(preferred: int = 8501, attempts: int = PORT_SEARCH_ATTEMPTS) -> int:
    for port in range(preferred, preferred + attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind((LOCAL_CONNECT_HOST, port))
                return port
            except OSError:
                continue
    raise RuntimeError("Kein freier lokaler Port fuer die App gefunden.")


def _remember_port(runtime: dict, port: int) -> None:
    try:
        port_file = runtime["user_root"] / "logs" / "server_port.txt"
        port_file.parent.mkdir(parents=True, exist_ok=True)
        port_file.write_text(str(port), encoding="utf-8")
    except Exception:
        pass


def run_streamlit_child() -> int:
    from desktop_runtime import app_script_path, configure_streamlit_runtime, prepare_runtime_environment, report_startup_failure

    runtime = None
    try:
        configure_streamlit_runtime()
        os.chdir(_current_dir())
        runtime = prepare_runtime_environment()
        _log_startup_event(runtime, "child runtime prepared")

        import_start = time.perf_counter()
        import streamlit.web.cli as stcli
        _log_startup_event(runtime, f"streamlit.web.cli imported after {time.perf_counter() - import_start:.3f}s")

        _apply_streamlit_options()
        port = _find_available_port()
        _remember_port(runtime, port)
        _log_startup_event(runtime, f"streamlit options applied; selected port={port}")

        sys.argv = [
            "streamlit",
            "run",
            str(app_script_path()),
            "--global.developmentMode=false",
            "--server.headless=true",
            f"--server.address={LOCAL_CONNECT_HOST}",
            f"--server.port={port}",
            "--server.showEmailPrompt=false",
            "--browser.gatherUsageStats=false",
        ]
        _log_startup_event(runtime, "calling streamlit cli main")
        return int(stcli.main() or 0)
    except Exception as exc:
        if runtime is not None:
            _log_startup_event(runtime, f"child failed: {exc}")
        report_startup_failure(exc)
        return 1


if __name__ == "__main__":
    if (
        os.environ.get(CHILD_ENV) == "1"
        or os.environ.get("EINSATZBERICHT_SUPPRESS_APP_SPLASH") == "1"
        or "--streamlit-child" in sys.argv
    ):
        sys.exit(run_streamlit_child())
    sys.exit(run_bootstrap())
