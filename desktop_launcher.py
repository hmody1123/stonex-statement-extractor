import subprocess
import sys
import webbrowser
import time
from pathlib import Path

APP_DIR = Path(__file__).parent
APP_FILE = APP_DIR / "app.py"

proc = subprocess.Popen([
    sys.executable, "-m", "streamlit", "run", str(APP_FILE),
    "--server.headless=true",
    "--server.port=8501",
    "--browser.gatherUsageStats=false",
])

time.sleep(5)
webbrowser.open("http://localhost:8501")

try:
    proc.wait()
except KeyboardInterrupt:
    proc.terminate()