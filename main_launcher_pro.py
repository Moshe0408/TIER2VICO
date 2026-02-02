import os
import sys
import subprocess
import threading
import queue
import json
import time
from datetime import datetime, timedelta
import io
import re
import shutil
import traceback


import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.widgets import ToastNotification

from tkinter.scrolledtext import ScrolledText
from tkinter import messagebox, StringVar
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

# ---------- CONFIG ----------
CONFIG_FILE = os.path.join(os.getcwd(), "config.json")

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "scripts": {},
        "theme": "darkly",
        "base_dir": os.getcwd()
    }

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=4, ensure_ascii=False)

# ---------- Unicode Fix ----------
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# ---------- Global ----------
CONFIG = load_config()
SCRIPTS = {
    "TIP": {
        "file": "TIP.PY",
        "cron": {"hour": 8, "minute": 0, "days": "sun,mon,tue,wed,thu"},
        "description": "דוח טיפים יומי (TIP)"
    },
    "STFPTEST": {
        "file": "STFPTEST.PY",
        "cron": {"hour": 5, "minute": 0},
        "description": "Stfp Mail Shufersal"
    },
    "STFPNOW": {
        "file": "STFPNOW.py",
        "cron": {"hour": 4, "minute": 30},
        "description": "STFP ורטיקלים"
    },
    "SHUFERSAL_GIFTCARD": {
        "file": "Shufersal_Giftcard.PY",
        "cron": {"hour": 12, "minute": 0},
        "description": "Shufersal Giftcard Report"
    },
    "TIER2": {
        "file": "TIER2.PY",
        "cron": {"hour": 8, "minute": 0},
        "description": "דוח ביצועים משולב (Tier 2)"
    },
    "DIGITAL": {
        "file": "Digital.py",
        "cron": {"hour": 8, "minute": 0},
        "description": "דוח דיגיטל (Digital)"
    }
}

BASE_DIR = CONFIG.get("base_dir", os.getcwd())

LOG_DIR = os.path.join(os.getcwd(), "logs")
os.makedirs(LOG_DIR, exist_ok=True)

log_q = queue.Queue()

# ---------- Helpers ----------
def timestamp():
    return datetime.now().strftime("[%Y-%m-%d %H:%M:%S] ")

def next_run_time(hour, minute):
    now = datetime.now()
    rt = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    if rt < now:
        rt += timedelta(days=1)
    return rt


# ---------- Script Runner ----------
def run_script_capture(name):
    import psutil
    info = SCRIPTS[name]
    
    # support internal running via strict args
    if info.get("internal") and info.get("arg"):
         # self execution
         cmd = [sys.executable, sys.argv[0], info["arg"]]
    else:
         script_path = os.path.join(BASE_DIR, info["file"])
         if not os.path.exists(script_path):
             log_q.put((name, f"{timestamp()}❌ File not found\n", "error", None))
             return
         cmd = [sys.executable, script_path]

    logfile = os.path.join(
        LOG_DIR, f"{name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    )

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    start = time.time()

    with open(logfile, "w", encoding="utf-8") as lf:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            env=env,
            errors="replace"
        )

        p = psutil.Process(proc.pid)
        log_q.put((name, "", "proc", p))
        log_q.put((name, f"{timestamp()}▶ Started {info['description']}\n", "info", None))

        for line in proc.stdout:
            lf.write(line)
            lf.flush()
            log_q.put((name, line, "log", None))

        proc.wait()
        duration = round(time.time() - start, 2)

        if proc.returncode == 0:
            log_q.put((name, f"{timestamp()}✔ Completed in {duration}s\n", "success", None))
        else:
            log_q.put((name, f"{timestamp()}❌ Failed ({proc.returncode})\n", "error", None))

def start_script_thread(name):
    threading.Thread(target=run_script_capture, args=(name,), daemon=True).start()

# ---------- Scheduler ----------
scheduler = BackgroundScheduler()

def schedule_jobs():
    scheduler.remove_all_jobs()
    for name, meta in SCRIPTS.items():
        cron = meta.get("cron")
        if cron:
            days = cron.get("days", "*") # Default to every day if not specified
            scheduler.add_job(
                lambda n=name: start_script_thread(n),
                CronTrigger(hour=cron["hour"], minute=cron["minute"], day_of_week=days),
                id=name,
                replace_existing=True,
                max_instances=1
            )
    if not scheduler.running:
        scheduler.start()

# ---------- JobCard ----------
class JobCard(ttk.Frame):
    def __init__(self, master, name, cfg, run_cb):
        super().__init__(master, padding=15, bootstyle="dark")
        self.name = name
        self.pack(fill="x", pady=8)

        # Glass effect frame
        self.card = ttk.Frame(self, padding=15, bootstyle="dark")
        self.card.pack(fill="both", expand=True)

        # Top row
        top = ttk.Frame(self.card)
        top.pack(fill="x")

        self.icon = ttk.Label(top, text="🕒", font=("Segoe UI", 18), foreground="#888888")
        self.icon.pack(side="left")

        ttk.Label(top, text=cfg["description"], font=("Segoe UI", 14, "bold")).pack(side="left", padx=10)

        self.badge = ttk.Label(top, text="Idle", bootstyle="secondary")
        self.badge.pack(side="right")

        # Progress
        self.progress = ttk.Progressbar(self.card, mode="determinate", bootstyle="info")
        self.progress.pack(fill="x", pady=5)

        # Bottom row
        bot = ttk.Frame(self.card)
        bot.pack(fill="x", pady=(5,0))

        self.run_btn = ttk.Button(bot, text="⚡ Run", bootstyle="success", command=lambda: run_cb(name))
        self.run_btn.pack(side="left")

        self.stop_btn = ttk.Button(bot, text="⏹ Stop", bootstyle="danger", state="disabled")
        self.stop_btn.pack(side="left", padx=5)

        self.next_run = StringVar(value="--:--")
        ttk.Label(bot, textvariable=self.next_run).pack(side="right")

    def set_status(self, txt, style):
        self.badge.config(text=txt, bootstyle=style)
        icon_map = {"Running":"⏳", "Success":"✔","Failed":"❌","Idle":"🕒"}
        color_map = {"Running":"#3498db","Success":"#00bc8c","Failed":"#cf6679","Idle":"#888888"}
        self.icon.config(text=icon_map.get(txt, "🕒"), foreground=color_map.get(txt, "#ffffff"))

    def start_progress(self, stop_cb=None):
        self.progress.start(10)
        self.stop_btn.config(state="normal", command=stop_cb)

    def stop_progress(self):
        self.progress.stop()
        self.stop_btn.config(state="disabled", command=None)

    def set_next_run(self, text):
        self.next_run.set(f"Next: {text}")

# ---------- Main App ----------
class LauncherApp(ttk.Window):
    def __init__(self):
        super().__init__(themename=CONFIG.get("theme", "darkly"))
        self.title("Moshe Automation Hub Pro (Integrated)")
        self.geometry("1400x850")

        self.job_cards = {}
        self.running_procs = {}

        self.build_ui()
        self.after(200, self.poll_logs)
        self.update_clock()
        self.update_countdowns()

    def build_ui(self):
        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)

        header = ttk.Frame(main)
        header.pack(fill="x")

        ttk.Label(header, text="Task Dashboard", font=("Segoe UI", 20, "bold")).pack(side="left")
        self.clock = ttk.Label(header, font=("Segoe UI", 16), foreground="#03dac6")
        self.clock.pack(side="right")

        # Dashboard tabs
        self.tabs = ttk.Notebook(main)
        self.tabs.pack(fill="both", expand=True, pady=10)

        # Dashboard tab with Scrollbar
        dash_container = ttk.Frame(self.tabs)
        self.tabs.add(dash_container, text="Dashboard")

        canvas = ttk.Canvas(dash_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(dash_container, orient="vertical", command=canvas.yview)
        self.dash_scroll_frame = ttk.Frame(canvas)

        self.dash_scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.dash_scroll_frame, anchor="nw", width=1350)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            
        # Bind safely
        try:
             canvas.bind_all("<MouseWheel>", _on_mousewheel)
        except: pass

        for name, cfg in SCRIPTS.items():
            card = JobCard(self.dash_scroll_frame, name, cfg, self.run_now)
            self.job_cards[name] = card

        logs = ttk.Frame(self.tabs)
        self.tabs.add(logs, text="Logs")

        self.log_area = ScrolledText(logs, font=("Consolas", 10), height=25)
        self.log_area.pack(fill="both", expand=True)
        self.log_area.tag_configure("success", foreground="#00bc8c")
        self.log_area.tag_configure("error", foreground="#e74c3c")
        self.log_area.tag_configure("info", foreground="#3498db")

    def run_now(self, name):
        card = self.job_cards[name]
        card.set_status("Running", "info")
        card.start_progress(lambda: self.stop_now(name))
        start_script_thread(name)

    def stop_now(self, name):
        proc = self.running_procs.get(name)
        if proc:
            try:
                for child in proc.children(recursive=True):
                    child.kill()
                proc.kill()
                self.log_area.insert("end", f"{timestamp()}🛑 {name} stopped by user\n", "error")
            except:
                self.log_area.insert("end", f"{timestamp()}🛑 Failed to stop {name}\n", "error")
        card = self.job_cards[name]
        card.stop_progress()
        card.set_status("Idle","secondary")

    def poll_logs(self):
        try:
            while True:
                name, msg, typ, data = log_q.get_nowait()

                if typ=="proc":
                    self.running_procs[name] = data
                elif typ=="log":
                    self.log_area.insert("end", f"[{name}] {msg}")
                elif typ=="success":
                    card = self.job_cards[name]
                    card.set_status("Success","success")
                    card.stop_progress()
                    self.log_area.insert("end", msg,"success")
                    if name in self.running_procs:
                        del self.running_procs[name]
                elif typ=="error":
                    card = self.job_cards[name]
                    card.set_status("Failed","danger")
                    card.stop_progress()
                    self.log_area.insert("end", msg,"error")
                    if name in self.running_procs:
                        del self.running_procs[name]
                else:
                    self.log_area.insert("end", msg)

                self.log_area.see("end")
        except queue.Empty:
            pass
        self.after(200,self.poll_logs)

    def update_clock(self):
        self.clock.config(text=datetime.now().strftime("%H:%M:%S"))
        self.after(1000,self.update_clock)

    def update_countdowns(self):
        now = datetime.now()
        for name in SCRIPTS:
            job = scheduler.get_job(name)
            if job and job.next_run_time:
                # Calculate delta using timezone naive objects for simplicity if needed, or keeping it aware
                # apscheduler uses tz aware datetimes usually
                nxt = job.next_run_time.replace(tzinfo=None)
                if nxt > now:
                     delta = nxt - now
                     self.job_cards[name].set_next_run(str(delta).split(".")[0])
                else:
                     self.job_cards[name].set_next_run("Running / Due")
            else:
                 self.job_cards[name].set_next_run("--:--")
        self.after(1000,self.update_countdowns)

# ---------- Entry ----------
if __name__ == "__main__":
    # Normal GUI Mode
    schedule_jobs()
    app = LauncherApp()
    app.mainloop()
