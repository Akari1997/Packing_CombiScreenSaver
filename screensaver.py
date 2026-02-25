import os
import sys
import time
import json
import random
import threading
import shutil
import tkinter as tk
from tkinter import Label
from PIL import Image, ImageTk
import pystray
from pystray import MenuItem as item
from PIL import Image as PILImage
import subprocess
from concurrent.futures import ThreadPoolExecutor
import ctypes
from ctypes import wintypes

# ================== 默认配置 ==================
DEFAULT_CONFIG = {
    "timeout": 120,
    "image_path": r"\\cnawww-s01\pub$\App_Packing\App\CombiScreenSaver\images",
    "interval": 3,
    "random_order": True,
    "tray_icon": "./icon.png",
    "cache_path": r"C:\Temp\images"
}

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")
CONFIG = DEFAULT_CONFIG.copy()

# ================== 读取配置 ==================
if not os.path.exists(CONFIG_FILE):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(CONFIG, f, indent=4, ensure_ascii=False)

try:
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        user_config = json.load(f) or {}
    for key, default_value in DEFAULT_CONFIG.items():
        CONFIG[key] = user_config.get(key, default_value)
except:
    CONFIG = DEFAULT_CONFIG.copy()

TIMEOUT = CONFIG["timeout"]
IMAGE_PATH = CONFIG["image_path"]
INTERVAL = CONFIG["interval"]
RANDOM_ORDER = CONFIG["random_order"]
ICON_PATH = CONFIG["tray_icon"]
CACHE_PATH = CONFIG["cache_path"]

if ICON_PATH and not os.path.isabs(ICON_PATH):
    ICON_PATH = os.path.join(os.path.dirname(__file__), ICON_PATH)

# ================== 全局状态 ==================
screensaver_running = False
screensaver_lock = threading.Lock()
root = None
sync_lock = threading.Lock()

# ================== 防止双开 ==================
def ensure_single_instance():
    mutex_name = "Global\\CombiScreenSaverMutex"
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
    last_error = ctypes.windll.kernel32.GetLastError()
    ERROR_ALREADY_EXISTS = 183
    if last_error == ERROR_ALREADY_EXISTS:
        sys.exit(0)

# ================== Windows 空闲检测 ==================
def get_idle_duration():
    class LASTINPUTINFO(ctypes.Structure):
        _fields_ = [
            ('cbSize', wintypes.UINT),
            ('dwTime', wintypes.DWORD),
        ]
    lii = LASTINPUTINFO()
    lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
    ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii))
    millis = ctypes.windll.kernel32.GetTickCount() - lii.dwTime
    return millis / 1000.0

# ================== 网络共享连接 ==================
def ensure_network_connection():
    try:
        if IMAGE_PATH.startswith("\\\\"):
            parts = IMAGE_PATH.split("\\")
            if len(parts) >= 4:
                server = parts[2]
                share = parts[3]
                unc_root = f"\\\\{server}\\{share}"
                subprocess.run(f'net use {unc_root} /persistent:no', shell=True)
    except:
        pass

# ================== 缓存同步 ==================
def ensure_cache_fast():
    with sync_lock:
        if not os.path.exists(CACHE_PATH):
            os.makedirs(CACHE_PATH)
        try:
            remote_files = {f: os.path.join(IMAGE_PATH, f)
                            for f in os.listdir(IMAGE_PATH)
                            if f.lower().endswith((".jpg", ".jpeg", ".png"))}
            local_files = {f: os.path.join(CACHE_PATH, f)
                           for f in os.listdir(CACHE_PATH)
                           if f.lower().endswith((".jpg", ".jpeg", ".png"))}
            # 删除多余
            for f in set(local_files) - set(remote_files):
                try:
                    os.remove(local_files[f])
                except:
                    pass
            def copy_file(f):
                try:
                    shutil.copy2(remote_files[f], os.path.join(CACHE_PATH, f))
                except:
                    pass
            with ThreadPoolExecutor(max_workers=5) as executor:
                for f in remote_files:
                    executor.submit(copy_file, f)
        except:
            pass

# ================== 图片缩放 ==================
def resize_contain(img, target_w, target_h):
    iw, ih = img.size
    scale = min(target_w / iw, target_h / ih)
    nw, nh = int(iw * scale), int(ih * scale)
    img = img.resize((nw, nh), Image.LANCZOS)
    bg = Image.new("RGB", (target_w, target_h), (0, 0, 0))
    x = (target_w - nw) // 2
    y = (target_h - nh) // 2
    bg.paste(img, (x, y))
    return bg

# ================== 屏保 ==================
def start_screensaver():
    global screensaver_running
    with screensaver_lock:
        if screensaver_running:
            return
        screensaver_running = True

    window = tk.Toplevel(root)
    window.configure(bg="black")
    window.overrideredirect(True)
    window.state("zoomed")
    window.attributes("-topmost", True)
    screen_w = window.winfo_screenwidth()
    screen_h = window.winfo_screenheight()
    label = Label(window, bg="black")
    label.pack(fill="both", expand=True)

    def exit_screensaver(event=None):
        global screensaver_running
        screensaver_running = False
        window.destroy()

    window.bind("<Motion>", exit_screensaver)
    window.bind("<Button>", exit_screensaver)
    window.bind("<Key>", exit_screensaver)
    window.focus_force()

    def show_loop():
        if not screensaver_running:
            return
        with sync_lock:
            paths = [os.path.join(CACHE_PATH, f)
                     for f in os.listdir(CACHE_PATH)
                     if f.lower().endswith((".jpg", ".jpeg", ".png"))]
        if not paths:
            window.after(3000, show_loop)
            return
        path = random.choice(paths) if RANDOM_ORDER else paths[int(time.time()) % len(paths)]
        try:
            img = Image.open(path)
            img = resize_contain(img, screen_w, screen_h)
            photo = ImageTk.PhotoImage(img)
            label.config(image=photo)
            label.image = photo
        except:
            pass
        window.after(INTERVAL * 1000, show_loop)

    show_loop()

# ================== 托盘 ==================
def tray_start(icon, item):
    root.after(0, start_screensaver)

def tray_exit(icon, item):
    icon.stop()
    root.quit()

def create_tray():
    if ICON_PATH and os.path.exists(ICON_PATH):
        image = PILImage.open(ICON_PATH).resize((32, 32))
    else:
        image = PILImage.new("RGB", (32, 32), (0, 0, 0))
    menu = (item("启动屏保", tray_start), item("退出", tray_exit))
    icon = pystray.Icon("ScreenSaver", image, "CombiScreenSaver", menu)
    icon.run_detached()

# ================== 监控 ==================
def monitor():
    if not screensaver_running and get_idle_duration() > TIMEOUT:
        start_screensaver()
    root.after(2000, monitor)

# ================== 主程序 ==================
if __name__ == "__main__":
    ensure_single_instance()  # ← 防止双开

    ensure_network_connection()
    ensure_cache_fast()

    root = tk.Tk()
    root.withdraw()

    threading.Thread(target=create_tray, daemon=True).start()
    root.after(2000, monitor)

    def background_sync():
        while True:
            time.sleep(60)
            ensure_cache_fast()

    threading.Thread(target=background_sync, daemon=True).start()

    root.mainloop()