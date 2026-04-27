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


def _open_browser_when_ready(port: int, timeout_seconds: int = 45) -> None:
    health_url = f"http://127.0.0.1:{port}/_stcore/health"
    browser_url = f"http://127.0.0.1:{port}"
    deadline = time.time() + timeout_seconds

    while time.time() < deadline:
        try:
            with urllib.request.urlopen(health_url, timeout=2):
                webbrowser.open(browser_url)
                return
        except Exception:
            time.sleep(0.5)


if __name__ == "__main__":
    try:
        current_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(__file__)
        os.chdir(current_dir)

        prepare_runtime_environment()
        _apply_streamlit_options()
        if maybe_check_for_updates():
            sys.exit(0)

        port = _find_available_port()
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
