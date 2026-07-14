import os
import sys
import subprocess
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CODE_DIR = PROJECT_ROOT / "code"
INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"

SCRAPER = CODE_DIR / "pipelinescrapper_mod_with_branch_canonical.py"
GUI_RUNNER = CODE_DIR / "gui_runner_with_branch_canonical_INTEGRATED.py"
DASHBOARD = CODE_DIR / "dashboard_app_with_branch_canonical_INTEGRATED.py"

REQUIRED_PACKAGES = [
    "pandas",
    "openpyxl",
    "playwright",
    "streamlit",
    "plotly",
    "beautifulsoup4"
]

def install_packages():
    print("\nInstalling required Python packages...\n")
    for package in REQUIRED_PACKAGES:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

    print("\nInstalling Playwright browsers...\n")
    subprocess.check_call([sys.executable, "-m", "playwright", "install"])

    print("\nAll dependencies installed successfully.\n")

def run_pipeline():
    print("\nRunning Job Scraper Pipeline...\n")
    subprocess.run([sys.executable, str(SCRAPER)], cwd=str(PROJECT_ROOT))

def run_dashboard():
    print("\nLaunching Streamlit Dashboard...\n")
    subprocess.run(
        [sys.executable, "-m", "streamlit", "run", str(DASHBOARD)],
        cwd=str(PROJECT_ROOT)
    )

def run_gui():
    print("\nLaunching GUI Runner...\n")
    subprocess.run([sys.executable, str(GUI_RUNNER)], cwd=str(PROJECT_ROOT))

def main():
    OUTPUT_DIR.mkdir(exist_ok=True)

    if len(sys.argv) == 1:
        install_packages()
        run_pipeline()

    elif sys.argv[1] == "gui":
        install_packages()
        run_gui()

    elif sys.argv[1] == "dash":
        install_packages()
        run_dashboard()

    else:
        print("""
Usage:

python code/setup_and_run.py
    Install packages and run scraper

python code/setup_and_run.py gui
    Launch GUI application

python code/setup_and_run.py dash
    Launch Streamlit dashboard
""")

if __name__ == "__main__":
    main()