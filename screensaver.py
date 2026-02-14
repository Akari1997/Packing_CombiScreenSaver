import os
import time
import json
import random
import threading
import shutil
import tkinter as tk
from tkinter import Label
from PIL import Image, ImageTk
from pynput import mouse, keyboard
import pystray
from pystray import MenuItem as item
from PIL import Image as PILImage
import subprocess
from concurrent.futures import ThreadPoolExecutor

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

if ICON_PATH:
    ICON_PATH = os.path.join(os.path.dirname(__file__), ICON_PATH)

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
    except Exception as e:
        print("网络连接异常:", e)

# ================== 全局状态 ==================
last_active_time = time.time()
screensaver_lock = threading.Lock()  # 避免多次启动冲突
stop_threads_flag = threading.Event()  # 控制异步线程停止

root = None

# ================== Pillow 兼容 ==================
try:
    RESAMPLE_MODE = Image.Resampling.LANCZOS
except:
    RESAMPLE_MODE = Image.ANTIALIAS

# ================== 工具函数 ==================
def resize_contain(img, target_w, target_h):
    iw, ih = img.size
    scale = min(target_w / iw, target_h / ih)
    nw, nh = int(iw * scale), int(ih * scale)
    img = img.resize((nw, nh), RESAMPLE_MODE)
    bg = Image.new("RGB", (target_w, target_h), (0, 0, 0))
    x = (target_w - nw) // 2
    y = (target_h - nh) // 2
    bg.paste(img, (x, y))
    return bg

# ================== 缓存同步 ==================
def ensure_cache_fast(strict_sync=False):
    if not os.path.exists(CACHE_PATH):
        os.makedirs(CACHE_PATH)
    try:
        remote_files = {f: os.path.join(IMAGE_PATH, f)
                        for f in os.listdir(IMAGE_PATH)
                        if f.lower().endswith((".jpg", ".jpeg", ".png"))}
        local_files = {f: os.path.join(CACHE_PATH, f)
                       for f in os.listdir(CACHE_PATH)
                       if f.lower().endswith((".jpg", ".jpeg", ".png"))}

        # 删除本地多余文件
        for f in set(local_files) - set(remote_files):
            try:
                os.remove(os.path.join(CACHE_PATH, f))
            except:
                pass

        # 首张立即同步
        first_file = next(iter(remote_files), None)
        if first_file:
            src = remote_files[first_file]
            dst = os.path.join(CACHE_PATH, first_file)
            if first_file not in local_files or os.path.getsize(dst) != os.path.getsize(src):
                shutil.copy2(src, dst)

        # 异步同步剩余图片
        def copy_file(f):
            try:
                shutil.copy2(remote_files[f], os.path.join(CACHE_PATH, f))
            except:
                pass
        with ThreadPoolExecutor(max_workers=5) as executor:
            for f in remote_files:
                if f != first_file:
                    executor.submit(copy_file, f)
    except Exception as e:
        print("缓存同步异常:", e)

# ================== 输入监听 ==================
def on_input(*args):
    global last_active_time
    last_active_time = time.time()
    stop_all_screensavers()

mouse.Listener(on_move=on_input, on_click=on_input, on_scroll=on_input).start()
keyboard.Listener(on_press=on_input).start()

# ================== 屏保管理 ==================
active_screensavers = []

def start_screensaver():
    """启动屏保，独立窗口+线程管理"""
    global active_screensavers, stop_threads_flag

    with screensaver_lock:
        stop_all_screensavers()
        stop_threads_flag.clear()

        # 局部状态
        images_cache = {}
        img_indices = []
        current_index = 0

        # 创建窗口
        window = tk.Toplevel(root)
        window.configure(bg="black")
        window.overrideredirect(True)
        window.state("zoomed")
        window.attributes("-topmost", True)
        window.bind("<Escape>", lambda e: stop_all_screensavers())

        label = Label(window, bg="gray", fg="white", text="加载中…", font=("Arial", 40))
        label.pack(fill="both", expand=True)

        screen_w = window.winfo_screenwidth()
        screen_h = window.winfo_screenheight()

        # 图片列表
        paths = [os.path.join(CACHE_PATH, f) for f in os.listdir(CACHE_PATH)
                 if f.lower().endswith((".jpg", ".jpeg", ".png"))]
        if not paths:
            return

        img_indices = list(range(len(paths)))
        if RANDOM_ORDER:
            random.shuffle(img_indices)
        current_index = 0

        # 加载首张图片
        try:
            img = Image.open(paths[0])
            img = resize_contain(img, screen_w, screen_h)
            images_cache[0] = ImageTk.PhotoImage(img)
            label.config(image=images_cache[0], text="")
            label.image = images_cache[0]
            current_index = 1
        except:
            pass

        # 异步加载剩余
        def load_remaining():
            for i in range(1, len(paths)):
                if stop_threads_flag.is_set():
                    break
                try:
                    img = Image.open(paths[i])
                    img = resize_contain(img, screen_w, screen_h)
                    images_cache[i] = ImageTk.PhotoImage(img)
                except:
                    pass
        threading.Thread(target=load_remaining, daemon=True).start()

        # 轮播
        def show_next():
            nonlocal current_index
            if stop_threads_flag.is_set() or not window.winfo_exists():
                return
            idx = img_indices[current_index]
            current_index = (current_index + 1) % len(img_indices)
            if idx in images_cache:
                label.config(image=images_cache[idx])
                label.image = images_cache[idx]
            window.after(INTERVAL*1000, show_next)

        show_next()
        active_screensavers.append(window)

def stop_all_screensavers():
    """关闭所有屏保窗口"""
    global active_screensavers, stop_threads_flag
    stop_threads_flag.set()
    for w in active_screensavers:
        try:
            if w.winfo_exists():
                w.destroy()
        except:
            pass
    active_screensavers = []

# ================== 托盘 ==================
def tray_start(icon, item):
    start_screensaver()

def tray_exit(icon, item):
    icon.stop()
    root.quit()

def create_tray():
    if ICON_PATH and os.path.exists(ICON_PATH):
        image = PILImage.open(ICON_PATH).resize((32,32))
    else:
        image = PILImage.new("RGB",(32,32),(0,0,0))
    menu = (item("启动屏保", tray_start),
            item("退出", tray_exit))
    icon = pystray.Icon("ScreenSaver", image, "包装屏保程序", menu)
    icon.run_detached()

# ================== 监控 ==================
def monitor():
    global last_active_time
    if time.time() - last_active_time > TIMEOUT:
        start_screensaver()
    root.after(1000, monitor)

# ================== 主程序 ==================
if __name__ == "__main__":
    ensure_network_connection()
    ensure_cache_fast(strict_sync=True)

    root = tk.Tk()
    root.withdraw()

    threading.Thread(target=create_tray, daemon=False).start()
    root.after(1000, monitor)

    # 后台增量同步
    def background_sync():
        while True:
            time.sleep(60)
            ensure_cache_fast(strict_sync=True)
    threading.Thread(target=background_sync, daemon=True).start()

    root.mainloop()