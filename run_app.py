"""
Entry point used to package the app into a standalone executable with PyInstaller.
Launches the Streamlit server programmatically and opens the browser, so the
recipient only has to double-click the .exe -- no Python install required.
"""
import os
import sys

from streamlit.web import cli as stcli


def resolve_path(path: str) -> str:
    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, path)


if __name__ == "__main__":
    os.chdir(resolve_path("."))
    sys.argv = [
        "streamlit",
        "run",
        resolve_path("app.py"),
        "--global.developmentMode=false",
    ]
    sys.exit(stcli.main())
