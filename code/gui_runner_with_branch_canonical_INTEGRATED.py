# gui_runner_with_branch_canonical_INTEGRATED.py
import os
import threading
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

# Resolve project folders
SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent
CODE_DIR = SCRIPT_DIR
INPUT_DIR = PROJECT_DIR / "input"
OUTPUT_DIR = PROJECT_DIR / "output"
os.chdir(CODE_DIR)


from pipelinescrapper_mod_with_branch_canonical import run_pipeline

COUNTS = ["50", "100", "150", "200", "250", "300", "Manual"]

DASHBOARD_APP = CODE_DIR / "dashboard_app_with_branch_canonical_INTEGRATED.py"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Jobright Pipeline GUI (Integrated Dashboard)")
        self.geometry("900x560")

        self.login_event = threading.Event()
        self.worker_thread = None

        # --- Top controls ---
        top = ttk.Frame(self)
        top.pack(fill="x", padx=12, pady=10)

        ttk.Label(top, text="Select job count:").pack(side="left")

        self.count_var = tk.StringVar(value="200")
        self.combo = ttk.Combobox(top, textvariable=self.count_var, values=COUNTS, width=12, state="readonly")
        self.combo.pack(side="left", padx=8)
        self.combo.bind("<<ComboboxSelected>>", self._toggle_manual)

        self.manual_var = tk.StringVar(value="")
        self.manual_entry = ttk.Entry(top, textvariable=self.manual_var, width=10)
        self.manual_entry.pack(side="left")
        self.manual_entry.configure(state="disabled")

        self.start_btn = ttk.Button(top, text="Start Scraping", command=self.start)
        self.start_btn.pack(side="left", padx=10)

        self.continue_btn = ttk.Button(top, text="Continue after login", command=self.cont_after_login, state="disabled")
        self.continue_btn.pack(side="left", padx=8)

        self.dashboard_btn = ttk.Button(top, text="Open Integrated Dashboard", command=self.open_dashboard)
        self.dashboard_btn.pack(side="left", padx=8)

        # --- Log box ---
        mid = ttk.Frame(self)
        mid.pack(fill="both", expand=True, padx=12, pady=10)

        ttk.Label(mid, text="Status / Logs:").pack(anchor="w")
        self.log = tk.Text(mid, height=22, wrap="word")
        self.log.pack(fill="both", expand=True)

        hint = (
            "Flow:\n"
            "1) Click Start Scraping → a browser opens.\n"
            "2) Login to Jobright in that browser.\n"
            "3) Come back and click 'Continue after login'.\n"
            "4) After finish, click 'Open Integrated Dashboard' (Main + Branch Summary).\n"
            "\n"
            "Inside the dashboard:\n"
            "- Main Dashboard tab shows branch metrics + multi-X chart.\n"
            "- Branch Salary Summary tab lets you generate the summary excel and see charts.\n"
        )
        self._append(hint)

    def _toggle_manual(self, _evt=None):
        is_manual = (self.count_var.get() == "Manual")
        self.manual_entry.configure(state=("normal" if is_manual else "disabled"))

    def _append(self, msg: str):
        self.log.insert("end", msg + "\n")
        self.log.see("end")

    def _get_target(self) -> int:
        sel = self.count_var.get()
        if sel == "Manual":
            try:
                n = int(self.manual_var.get().strip())
            except Exception:
                raise ValueError("Manual entry must be a number.")
            if n <= 0:
                raise ValueError("Manual entry must be > 0.")
            return n
        return int(sel)

    def cont_after_login(self):
        self.login_event.set()
        self.continue_btn.configure(state="disabled")
        self._append("[GUI] Continue clicked. Scraper will proceed...")

    def start(self):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("Running", "Scraper is already running.")
            return

        try:
            target = self._get_target()
        except ValueError as e:
            messagebox.showerror("Invalid Input", str(e))
            return

        self.login_event.clear()
        self.start_btn.configure(state="disabled")
        self.continue_btn.configure(state="normal")
        self._append(f"[GUI] Starting pipeline for target_jobs={target}")

        def wait_for_login():
            self._append("[GUI] Waiting for 'Continue after login'...")
            self.login_event.wait()

        def on_status(msg: str):
            self.after(0, lambda: self._append(msg))

        def worker():
            try:
                run_pipeline(
                    target_jobs=target,
                    headless=False,
                    wait_for_login=wait_for_login,
                    on_status=on_status,
                )
                self.after(0, lambda: self._append(f"[GUI] Pipeline completed. Excel ready: {OUTPUT_DIR / 'jobright_jobs.xlsx'}"))
            except Exception as e:
                self.after(0, lambda: self._append(f"[GUI] ERROR: {e}"))
            finally:
                self.after(0, lambda: self.start_btn.configure(state="normal"))
                self.after(0, lambda: self.continue_btn.configure(state="disabled"))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    def open_dashboard(self):
        self._append("[GUI] Launching integrated dashboard (Streamlit)...")
        app = Path(DASHBOARD_APP)
        if not app.exists():
            messagebox.showerror("Missing file", f"Missing dashboard file: {DASHBOARD_APP}")
            return
        try:
            subprocess.Popen([sys.executable, "-m", "streamlit", "run", str(DASHBOARD_APP)], cwd=str(PROJECT_DIR))
        except Exception as e:
            messagebox.showerror("Dashboard Error", str(e))


if __name__ == "__main__":
    App().mainloop()
