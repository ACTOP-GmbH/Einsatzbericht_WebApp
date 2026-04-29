import os
import socket
import sys
import threading
import time
import urllib.request
import webbrowser

from desktop_runtime import (
    app_script_path,
    configure_streamlit_runtime,
    maybe_check_for_updates,
    prepare_runtime_environment,
    report_startup_failure,
)

configure_streamlit_runtime()

import streamlit.web.cli as stcli
from streamlit import config as st_config


def _apply_streamlit_options() -> None:
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


def _find_available_port(preferred: int = 8501, attempts: int = 20) -> int:
    for port in range(preferred, preferred + attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("Kein freier lokaler Port fuer die App gefunden.")


def _server_ready(port: int, timeout: float = 1.0) -> bool:
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/_stcore/health", timeout=timeout):
            return True
    except Exception:
        return False


def _open_existing_instance(runtime: dict, preferred: int = 8501, attempts: int = 20) -> bool:
    port_file = runtime["user_root"] / "logs" / "server_port.txt"
    ports = []
    if port_file.exists():
        try:
            ports.append(int(port_file.read_text(encoding="utf-8").strip()))
        except Exception:
            pass
    ports.extend(range(preferred, preferred + attempts))
    for port in dict.fromkeys(ports):
        if _server_ready(int(port)):
            webbrowser.open(f"http://127.0.0.1:{int(port)}")
            return True
    return False


def _remember_port(runtime: dict, port: int) -> None:
    try:
        port_file = runtime["user_root"] / "logs" / "server_port.txt"
        port_file.parent.mkdir(parents=True, exist_ok=True)
        port_file.write_text(str(port), encoding="utf-8")
    except Exception:
        pass


def _open_browser_when_ready(port: int, timeout_seconds: int = 45) -> None:
    browser_url = f"http://127.0.0.1:{port}"
    deadline = time.time() + timeout_seconds

    while time.time() < deadline:
        if _server_ready(port, timeout=2):
            webbrowser.open(browser_url)
            return
        time.sleep(0.5)


if __name__ == "__main__":
    try:
        current_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(__file__)
        os.chdir(current_dir)

        runtime = prepare_runtime_environment()
        _apply_streamlit_options()
        if maybe_check_for_updates():
            sys.exit(0)
        if _open_existing_instance(runtime):
            sys.exit(0)

        port = _find_available_port()
        _remember_port(runtime, port)
        threading.Thread(target=_open_browser_when_ready, args=(port,), daemon=True).start()

        sys.argv = [
            "streamlit",
            "run",
            str(app_script_path()),
            "--global.developmentMode=false",
            "--server.headless=true",
            "--server.address=127.0.0.1",
            f"--server.port={port}",
            "--server.showEmailPrompt=false",
            "--browser.gatherUsageStats=false",
        ]
        sys.exit(stcli.main())
    except Exception as exc:
        report_startup_failure(exc)
        sys.exit(1)
