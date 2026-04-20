import os
import sys

import streamlit.web.cli as stcli

from desktop_runtime import app_script_path, maybe_check_for_updates, prepare_runtime_environment


if __name__ == "__main__":
    current_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(__file__)
    os.chdir(current_dir)

    prepare_runtime_environment()
    if maybe_check_for_updates():
        sys.exit(0)

    sys.argv = [
        "streamlit",
        "run",
        str(app_script_path()),
        "--global.developmentMode=false",
    ]
    sys.exit(stcli.main())
