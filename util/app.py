import os
import sys
import time
from datetime import datetime

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


def parse_token_expire(login_data: dict):
    """
    计算 token 过期时间

    参数:
        login_data: 登录返回的 json

    返回:
        {
            "expire_timestamp": int,
            "expire_time": "YYYY-MM-DD HH:MM:SS",
            "seconds_left": int,
            "expired": bool
        }
    """

    expire_ts = login_data.get("expire")

    if not expire_ts:
        raise ValueError("json 中没有 expire 字段")

    now_ts = int(time.time())

    expired = now_ts >= expire_ts

    seconds_left = max(0, expire_ts - now_ts)

    expire_time_str = datetime.fromtimestamp(expire_ts).strftime("%Y-%m-%d %H:%M:%S")

    return {
        "expire_timestamp": expire_ts,
        "expire_time": expire_time_str,
        "seconds_left": seconds_left,
        "expired": expired,
    }