import os
import sys
import time

import pyautogui

from util import system


def get_version():
    try:
        with open(os.path.join(system.get_exe_dir(), "version.txt"), "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return ""


def press_any_key_exit():
    input("按 Enter (回车) 键退出程序...")
    sys.exit()


def ensure_start_by_self():
    if "run" not in sys.argv[1:]:
        sys.exit(0)


def check_load_finished(delay=3.0):
    pixel_color = (10, 90, 160)
    last_time = time.time()

    while True:
        if any(pyautogui.pixel(x, y) == pixel_color for x, y in [(992, 598), (993, 593)]):
            last_time = time.time()
        if time.time() - last_time >= delay:
            break
