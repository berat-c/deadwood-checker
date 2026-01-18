import os
import sys
import json
import time
import threading
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
import webbrowser
import psutil
import requests
import traceback
from PIL import Image, ImageDraw, ImageTk


# Tray
import pystray

# Startup registry (Windows)
import winreg

# WinAPI window title scan
import ctypes
from ctypes import wintypes


# ====== BACKEND CONFIG ======
PROCESS_NAME = "RedM_GTAProcess.exe"

CHECK_IDLE_SEC = 5        # slower loop when idle / not in deadwood
CHECK_ACTIVE_SEC = 2      # faster loop while confirming deadwood
REQUIRED_HITS = 2         # must see Deadwood this many consecutive checks
GRACE_AFTER_PROCESS_START_SEC = 30  # wait after RedM starts before title checks

# Webhook is handled on "backend" (not user-editable in UI)
WEBHOOK_URL = "https://discord.com/api/webhooks/1462018244432105626/gdy8dwYebfdUKhIpnKuYhMWseh4XIoLLizxP-Cl54Chb1mQdOiTlxFvy7EGPuzW-5TND"
# ============================

APP_NAME = "Deadwood Presence Checker"
RUN_KEY_NAME = "DeadwoodPresenceChecker"

APPDATA_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / APP_NAME
CONFIG_PATH = APPDATA_DIR / "config.json"
LOG_PATH = APPDATA_DIR / "log.txt"
APP_VERSION = "v0.5"


# ===== WinAPI: enumerate visible windows and read titles =====
user32 = ctypes.windll.user32

EnumWindows = user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
GetWindowTextLengthW = user32.GetWindowTextLengthW
GetWindowTextW = user32.GetWindowTextW
IsWindowVisible = user32.IsWindowVisible

# Safer signatures
EnumWindows.restype = wintypes.BOOL
EnumWindows.argtypes = [EnumWindowsProc, wintypes.LPARAM]

GetWindowTextLengthW.restype = ctypes.c_int
GetWindowTextLengthW.argtypes = [wintypes.HWND]

GetWindowTextW.restype = ctypes.c_int
GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]

IsWindowVisible.restype = wintypes.BOOL
IsWindowVisible.argtypes = [wintypes.HWND]


def any_window_title_contains(substring: str) -> bool:
    """
    Returns True if ANY visible top-level window title contains substring (case-insensitive).
    Stops early as soon as it finds a match.
    """
    target = substring.lower()
    found = False

    @EnumWindowsProc
    def enum_proc(hwnd, lparam):
        nonlocal found
        if found:
            return False  # stop early

        if not IsWindowVisible(hwnd):
            return True

        length = GetWindowTextLengthW(hwnd)
        if length <= 0:
            return True

        buf = ctypes.create_unicode_buffer(length + 1)
        GetWindowTextW(hwnd, buf, length + 1)

        title = (buf.value or "").strip().lower()
        if title and (target in title):
            found = True
            return False  # stop early

        return True

    EnumWindows(enum_proc, 0)
    return found


def is_process_running(proc_name: str) -> bool:
    target = proc_name.lower()
    for p in psutil.process_iter(["name"]):
        try:
            name = (p.info.get("name") or "").lower()
            if name == target:
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
        except Exception:
            continue
    return False


def send_webhook_message(content: str) -> None:
    log(f"Webhook: sending: {content}")
    try:
        r = requests.post(WEBHOOK_URL, json={"content": content}, timeout=10)
        r.raise_for_status()
        log(f"Webhook: sent OK (status={r.status_code})")
    except Exception as e:
        log(f"Webhook: FAILED: {e}\n{traceback.format_exc()}")
        raise


def ensure_config_dir():
    APPDATA_DIR.mkdir(parents=True, exist_ok=True)


def log(msg: str):
    """Append a timestamped line to %APPDATA%\Deadwood Presence Checker\log.txt. Never raises."""
    try:
        ensure_config_dir()
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass

def resource_path(relative_path: str) -> str:
    """
    Get absolute path to resource, works for dev and for PyInstaller one-file exe.
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def set_window_icon(root):
    try:
        img_path = resource_path("HavenBornLogo.png")
        img = Image.open(img_path).convert("RGBA")
        icon = ImageTk.PhotoImage(img)
        root.iconphoto(True, icon)
        root._icon_ref = icon  # prevent garbage collection
    except Exception as e:
        log(f"Failed to set window icon: {e}")


def load_config() -> dict:
    ensure_config_dir()
    if not CONFIG_PATH.exists():
        return {
            "nickname": "Ezekiel",
            "run_at_startup": False,
            "run_minimized": False,
            "start_monitoring_automatically": False,
            "always_notify": False,
        }
    try:
        cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        cfg = {}

    # Backward-compatible defaults
    cfg.setdefault("nickname", "Ezekiel")
    cfg.setdefault("run_at_startup", False)
    cfg.setdefault("run_minimized", False)
    cfg.setdefault("start_monitoring_automatically", False)
    cfg.setdefault("always_notify", False)
    return cfg


def save_config(cfg: dict) -> None:
    ensure_config_dir()
    CONFIG_PATH.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")


def get_startup_command() -> str:
    exe = sys.executable

    # If running as a .py script, prefer pythonw.exe to avoid a console window
    if exe.lower().endswith("python.exe"):
        pythonw = exe[:-10] + "pythonw.exe"
        if os.path.exists(pythonw):
            exe = pythonw

    script = os.path.abspath(sys.argv[0])

    # If packaged, sys.argv[0] is usually the exe itself
    if exe.lower().endswith(".exe") and script.lower().endswith(".exe"):
        return f'"{exe}"'
    return f'"{exe}" "{script}"'


def set_run_at_startup(enabled: bool) -> None:
    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\Run",
        0,
        winreg.KEY_SET_VALUE,
    ) as key:
        if enabled:
            winreg.SetValueEx(key, RUN_KEY_NAME, 0, winreg.REG_SZ, get_startup_command())
        else:
            try:
                winreg.DeleteValue(key, RUN_KEY_NAME)
            except FileNotFoundError:
                pass


def is_startup_enabled() -> bool:
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_READ,
        ) as key:
            winreg.QueryValueEx(key, RUN_KEY_NAME)
            return True
    except FileNotFoundError:
        return False
    except OSError:
        return False


def ask_user_to_announce(nickname: str) -> bool:
    temp = tk.Tk()
    temp.withdraw()
    temp.attributes("-topmost", True)
    msg = f'Seems like you are waking up as "{nickname}" in Deadwood.\nDo you wanna let people know?'
    res = messagebox.askyesno("Deadwood Presence", msg, parent=temp)
    temp.destroy()
    return res


def create_tray_icon_image() -> Image.Image:
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((6, 6, size - 6, size - 6), radius=14, outline=(255, 255, 255, 255), width=3)
    d.text((22, 16), "D", fill=(255, 255, 255, 255))
    d.ellipse((42, 40, 52, 50), fill=(255, 255, 255, 255))
    return img


class DeadwoodApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        self.root.resizable(False, False)

        self.cfg = load_config()

        # Monitoring state
        self.monitoring = False
        self.stop_event = threading.Event()
        self.worker_thread = None

        # Tray
        self.tray_icon = None
        self.tray_thread = None
        self.is_hidden_to_tray = False

        # UI variables
        self.nickname_var = tk.StringVar(value=self.cfg.get("nickname", "Ezekiel"))
        self.run_minimized_var = tk.BooleanVar(value=bool(self.cfg.get("run_minimized", False)))

        startup_state = is_startup_enabled() if self.cfg.get("run_at_startup") else False
        self.run_startup_var = tk.BooleanVar(value=startup_state)

        self.auto_monitor_var = tk.BooleanVar(value=bool(self.cfg.get("start_monitoring_automatically", False)))
        self.always_notify_var = tk.BooleanVar(value=bool(self.cfg.get("always_notify", False)))

        self.status_var = tk.StringVar(value=f"Status: Idle (watching {PROCESS_NAME})")
        self.build_ui()

        # Apply startup if config wants it but registry differs
        if self.cfg.get("run_at_startup", False) and not startup_state:
            try:
                set_run_at_startup(True)
                self.run_startup_var.set(True)
            except Exception:
                self.run_startup_var.set(False)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Auto start monitoring if enabled
        if self.auto_monitor_var.get():
            self.start_monitoring(minimize=self.run_minimized_var.get())

    def build_ui(self):
        pad = 12
        frame = tk.Frame(self.root, padx=pad, pady=pad)
        frame.pack()

        tk.Label(frame, text="In character name:").grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.nickname_var, width=32).grid(row=0, column=1, sticky="w", padx=(8, 0))

        # Checkboxes
        self.cb_startup = tk.Checkbutton(
            frame,
            text="Run at startup",
            variable=self.run_startup_var,
            command=self.on_toggle_startup,
        )
        self.cb_startup.grid(row=1, column=0, columnspan=2, sticky="w", pady=(10, 0))

        self.cb_minimized = tk.Checkbutton(
            frame,
            text="Run minimized",
            variable=self.run_minimized_var,
            command=self.on_toggle_any_setting,
        )
        self.cb_minimized.grid(row=2, column=0, columnspan=2, sticky="w")

        self.cb_auto_monitor = tk.Checkbutton(
            frame,
            text="Start monitoring automatically",
            variable=self.auto_monitor_var,
            command=self.on_toggle_any_setting,
        )
        self.cb_auto_monitor.grid(row=3, column=0, columnspan=2, sticky="w")

        self.cb_always_notify = tk.Checkbutton(
            frame,
            text="Always notify",
            variable=self.always_notify_var,
            command=self.on_toggle_any_setting,
        )
        self.cb_always_notify.grid(row=4, column=0, columnspan=2, sticky="w")

        tk.Label(frame, textvariable=self.status_var).grid(row=5, column=0, columnspan=2, sticky="w", pady=(10, 0))

        btns = tk.Frame(frame)
        btns.grid(row=6, column=0, columnspan=2, sticky="w", pady=(12, 0))

        self.btn_start = tk.Button(btns, text="Start monitoring", command=lambda: self.start_monitoring(minimize=False))
        self.btn_start.pack(side="left")

        self.btn_start_min = tk.Button(
            btns,
            text="Start monitoring and minimize",
            command=lambda: self.start_monitoring(minimize=True),
        )
        self.btn_start_min.pack(side="left", padx=(8, 0))

        self.btn_stop = tk.Button(btns, text="Stop", command=self.stop_monitoring, state="disabled")
        self.btn_stop.pack(side="left", padx=(8, 0))

        tk.Label(
            frame,
            text="- When minimized, use the tray icon menu to show / stop / exit.",
            fg="gray",
        ).grid(row=7, column=0, columnspan=2, sticky="w", pady=(10, 0))

        tk.Label(
            frame,
            text="- Closing this window will hide the app to system tray menu.",
            fg="gray",
        ).grid(row=8, column=0, columnspan=2, sticky="w")

        def open_github(event=None):
            webbrowser.open_new("https://github.com/berat-c/deadwood-checker")

        link = tk.Label(
            frame,
            text="Made by Biretro",
            fg="#4ea3ff",  # link-like blue
            cursor="hand2",
        )

        link.bind("<Button-1>", open_github)
        link.bind("<Enter>", lambda e: link.config(font=("Segoe UI", 9, "underline")))
        link.bind("<Leave>", lambda e: link.config(font=("Segoe UI", 9)))
        link.grid(row=9, column=0, columnspan=2, sticky="w", pady=(0, 4))

        tk.Label(
            frame,
            text=APP_VERSION,
            fg="gray",
            font=("Segoe UI", 8),  # smaller than everything else
        ).grid(row=10, column=0, columnspan=2, sticky="w", pady=(0, 4))

    def set_status(self, text: str):
        self.status_var.set(text)

    def persist_config(self):
        self.cfg["nickname"] = self.nickname_var.get().strip() or "Ezekiel"
        self.cfg["run_at_startup"] = bool(self.run_startup_var.get())
        self.cfg["run_minimized"] = bool(self.run_minimized_var.get())
        self.cfg["start_monitoring_automatically"] = bool(self.auto_monitor_var.get())
        self.cfg["always_notify"] = bool(self.always_notify_var.get())
        save_config(self.cfg)

    def on_toggle_any_setting(self):
        self.persist_config()
        self.set_status("Status: Saved")

    def on_toggle_startup(self):
        enabled = bool(self.run_startup_var.get())
        try:
            set_run_at_startup(enabled)
            self.persist_config()
            self.set_status("Status: Saved")
        except Exception as e:
            self.run_startup_var.set(not enabled)
            messagebox.showerror("Startup error", f"Couldn't update startup setting:\n{e}")
            self.set_status("Status: Failed to update startup")

    def start_monitoring(self, minimize: bool):
        nickname = self.nickname_var.get().strip()
        if not nickname:
            messagebox.showerror("Error", "Nickname cannot be empty.")
            return

        if self.monitoring:
            if minimize:
                self.minimize_to_tray()
            return

        self.persist_config()

        self.monitoring = True
        self.stop_event.clear()

        log("Monitoring started")

        self.btn_start.config(state="disabled")
        self.btn_start_min.config(state="disabled")
        self.btn_stop.config(state="normal")

        self.set_status(f"Status: Monitoring for RedM")

        self.worker_thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.worker_thread.start()

        if minimize:
            self.minimize_to_tray()

    def stop_monitoring(self):
        if not self.monitoring:
            return

        log("Monitoring stopped")
        self.stop_event.set()
        self.monitoring = False

        log("Monitoring stopped by user")

        self.btn_start.config(state="normal")
        self.btn_start_min.config(state="normal")
        self.btn_stop.config(state="disabled")

        self.set_status("Status: Stopped")

    def monitor_loop(self):
        was_running = False
        presence_announced = False
        presence_decided = False  # user has made a yes/no decision this session

        deadwood_hits = 0
        first_seen_running_ts = None

        while not self.stop_event.is_set():
            nickname = (self.nickname_var.get().strip() or "Ezekiel")
            always_notify = bool(self.always_notify_var.get())

            # Cheapest check first
            running = is_process_running(PROCESS_NAME)

            now = time.time()
            if running and not was_running:
                first_seen_running_ts = now
            if not running:
                first_seen_running_ts = None

            sleep_for = CHECK_IDLE_SEC
            in_deadwood_raw = False

            # Only after grace: check if any window title contains "Deadwood County"
            if running and first_seen_running_ts is not None:
                if (now - first_seen_running_ts) >= GRACE_AFTER_PROCESS_START_SEC:
                    try:
                        in_deadwood_raw = any_window_title_contains("Deadwood County")
                    except Exception:
                        in_deadwood_raw = False

            if running and in_deadwood_raw:
                deadwood_hits += 1
                sleep_for = CHECK_ACTIVE_SEC
            else:
                deadwood_hits = 0

            in_deadwood_now = (deadwood_hits >= REQUIRED_HITS)

            # Enter Deadwood (stable)
            if in_deadwood_now and not presence_decided:
                try:
                    if always_notify:
                        yes = True
                    else:
                        yes = ask_user_to_announce(nickname)

                    presence_decided = True  # IMPORTANT: latch decision (Yes OR No)

                    if yes:
                        send_webhook_message(f" :inbox_tray: **{nickname}** is around.")
                        presence_announced = True
                except Exception:
                    # If something fails, do NOT lock the user out forever.
                    # Only mark decided if we successfully got a decision.
                    pass

            # Game closed
            if (not running) and was_running:
                if presence_announced:
                    try:
                        send_webhook_message(f" :bed: **{nickname}** went to bed.")
                    except Exception:
                        pass

                # Reset session state ONLY when game closes
                presence_decided = False
                presence_announced = False

            was_running = running
            was_in_deadwood = in_deadwood_now

            # Sleep in small chunks so Stop is responsive
            waited = 0.0
            while waited < sleep_for and not self.stop_event.is_set():
                time.sleep(0.2)
                waited += 0.2

    # ===== Tray behavior =====
    def ensure_tray(self):
        if self.tray_icon is not None:
            return

        image = create_tray_icon_image()

        def on_show(icon, item):
            self.root.after(0, self.show_window)

        def on_stop(icon, item):
            self.root.after(0, self.stop_monitoring)

        def on_exit(icon, item):
            self.root.after(0, self.exit_app)

        menu = pystray.Menu(
            pystray.MenuItem("Open", on_show, default=True),  # <-- double-click triggers this
            pystray.MenuItem("Stop monitoring", on_stop),
            pystray.MenuItem("Exit", on_exit),
        )

        # Create icon ONCE; Windows double-click triggers the default menu item above.
        self.tray_icon = pystray.Icon(APP_NAME, image, APP_NAME, menu)

        def run_tray():
            self.tray_icon.run()

        self.tray_thread = threading.Thread(target=run_tray, daemon=True)
        self.tray_thread.start()

    def minimize_to_tray(self):
        self.ensure_tray()
        self.is_hidden_to_tray = True
        self.root.withdraw()
        self.set_status(f"Status: Monitoring in tray for RedM")

    def show_window(self):
        self.is_hidden_to_tray = False
        self.root.deiconify()
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(200, lambda: self.root.attributes("-topmost", False))

    def exit_app(self):
        try:
            self.stop_event.set()
            self.monitoring = False
        finally:
            if self.tray_icon is not None:
                try:
                    self.tray_icon.stop()
                except Exception:
                    pass
            self.persist_config()
            self.root.destroy()

    def on_close(self):
        # close button minimizes to tray
        self.minimize_to_tray()


def main():
    log("Application starting")
    root = tk.Tk()
    set_window_icon(root)   # 👈 THIS sets the feather icon
    app = DeadwoodApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
