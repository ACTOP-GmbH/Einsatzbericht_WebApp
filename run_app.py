import os
import sys

from desktop_launcher import CHILD_ENV, run_bootstrap, run_streamlit_child


if __name__ == "__main__":
    if (
        os.environ.get(CHILD_ENV) == "1"
        or os.environ.get("EINSATZBERICHT_SUPPRESS_APP_SPLASH") == "1"
        or "--streamlit-child" in sys.argv
    ):
        sys.exit(run_streamlit_child())
    sys.exit(run_bootstrap())
