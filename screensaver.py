import os
import time
import json
import random
import threading
import sys
import tkinter as tk
from tkinter import Label
from PIL import Image, ImageTk
from pynput import mouse, keyboard
import pystray
from pystray import MenuItem as item
from PIL import Image as PILImage
import winreg

import os
import json

# ================== 配置 ==================
DEFAULT_CONFIG = {
    "timeout": 120,
    "image_path": "./images",
    "interval": 3,
    "random_order": True,
    "tray_icon": "./icon.png"
}

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")

# 初始化 CONFIG
CONFIG = DEFAULT_CONFIG.copy()

# 如果文件不存在，生成完整配置
if not os.path.exists(CONFIG_FILE):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(CONFIG, f, indent=4, ensure_ascii=False)
    print("config.json 不存在，已生成默认配置文件")

# 尝试读取用户配置
try:
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        user_config = json.load(f) or {}

    # 合并用户配置和默认值，保证每个字段都有值
    for key, default_value in DEFAULT_CONFIG.items():
        CONFIG[key] = user_config.get(key, default_value)

except json.JSONDecodeError:
    print("config.json 格式错误，使用默认配置")
    CONFIG = DEFAULT_CONFIG.copy()

except Exception as e:
    print(f"读取配置文件失败：{e}，使用默认配置")
    CONFIG = DEFAULT_CONFIG.copy()

# 读取配置
TIMEOUT = CONFIG["timeout"]
IMAGE_PATH = CONFIG["image_path"]
INTERVAL = CONFIG["interval"]
RANDOM_ORDER = CONFIG["random_order"]
ICON_PATH = CONFIG["tray_icon"]

if ICON_PATH:
    ICON_PATH = os.path.join(os.path.dirname(__file__), ICON_PATH)

# 如果你想让第一次生成的 JSON 永远包含所有字段（包括缺省值），可以再写一次：
with open(CONFIG_FILE, "w", encoding="utf-8") as f:
    json.dump(CONFIG, f, indent=4, ensure_ascii=False)


# ================== 全局状态 ==================
last_active_time = time.time()
screensaver_active = False
start_request = False
stop_event = threading.Event()

root = None
screensaver_window = None
label = None
after_job = None
images = []
img_indices = []
current_index = 0

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


# ================== 输入监听 ==================
def on_input(*args):
    global last_active_time
    last_active_time = time.time()
    if screensaver_active:
        stop_screensaver()


mouse.Listener(on_move=on_input, on_click=on_input, on_scroll=on_input).start()
keyboard.Listener(on_press=on_input).start()


# ================== 屏保 ==================
def start_screensaver():
    global screensaver_active, screensaver_window
    global images, img_indices, current_index, label

    if screensaver_active:
        return

    screensaver_active = True

    screensaver_window = tk.Toplevel(root)
    screensaver_window.configure(bg="black")
    screensaver_window.overrideredirect(True)
    screensaver_window.state("zoomed")
    screensaver_window.attributes("-topmost", True)

    screensaver_window.bind("<Escape>", lambda e: stop_screensaver())

    images = [os.path.join(IMAGE_PATH, f)
              for f in os.listdir(IMAGE_PATH)
              if f.lower().endswith((".jpg", ".jpeg", ".png"))]

    if not images:
        stop_screensaver()
        return

    img_indices = list(range(len(images)))
    if RANDOM_ORDER:
        random.shuffle(img_indices)

    current_index = 0

    label = Label(screensaver_window, bg="black")
    label.pack(fill="both", expand=True)

    show_next()


def show_next():
    global current_index, after_job

    if not screensaver_active:
        return

    screen_w = screensaver_window.winfo_screenwidth()
    screen_h = screensaver_window.winfo_screenheight()

    idx = img_indices[current_index]
    current_index = (current_index + 1) % len(img_indices)

    img = Image.open(images[idx])
    img = resize_contain(img, screen_w, screen_h)
    photo = ImageTk.PhotoImage(img)

    label.config(image=photo)
    label.image = photo

    after_job = screensaver_window.after(INTERVAL * 1000, show_next)


def stop_screensaver():
    global screensaver_active, screensaver_window, after_job

    if not screensaver_active:
        return

    screensaver_active = False

    if after_job:
        try:
            screensaver_window.after_cancel(after_job)
        except:
            pass
        after_job = None

    if screensaver_window:
        try:
            screensaver_window.destroy()
        except:
            pass
        screensaver_window = None


# ================== 托盘 ==================
def tray_start(icon, item):
    global start_request
    start_request = True


def tray_exit(icon, item):
    stop_event.set()
    icon.stop()
    root.quit()


def create_tray():
    if ICON_PATH and os.path.exists(ICON_PATH):
        image = PILImage.open(ICON_PATH).resize((32, 32), RESAMPLE_MODE)
    else:
        image = PILImage.new("RGB", (32, 32), (0, 0, 0))

    menu = (item("启动屏保", tray_start),
            item("退出", tray_exit))

    icon = pystray.Icon("ScreenSaver", image, "包装屏保程序", menu)
    icon.run()


# ================== 监控 ==================
def monitor():
    global start_request

    if start_request:
        start_request = False
        start_screensaver()

    if not screensaver_active and (time.time() - last_active_time > TIMEOUT):
        start_screensaver()

    root.after(1000, monitor)


# ================== 注册开机启动 ==================
def register_startup():
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0, winreg.KEY_SET_VALUE)

        winreg.SetValueEx(
            key,
            "ScreenSaverTray",
            0,
            winreg.REG_SZ,
            sys.executable + " " + os.path.abspath(__file__))

        winreg.CloseKey(key)
    except Exception as e:
        print("注册开机启动失败:", e)


# ================== 主程序 ==================
if __name__ == "__main__":
    time.sleep(0)
    register_startup()

    root = tk.Tk()
    root.withdraw()  # 主窗口隐藏

    threading.Thread(target=create_tray, daemon=True).start()

    root.after(1000, monitor)
    root.mainloop()
