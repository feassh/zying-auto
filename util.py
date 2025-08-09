import ctypes
import json
import os
import platform
import sys

import cv2
import numpy as np
import psutil
import requests
from PIL import ImageGrab


def get_exe_dir():
    # PyInstaller打包后的程序，sys.executable 是exe路径
    # 脚本运行时，__file__ 是脚本路径
    if getattr(sys, 'frozen', False):
        # 打包后，exe模式
        return os.path.dirname(sys.executable)
    else:
        # 普通脚本模式
        return os.path.dirname(os.path.abspath(__file__))


config_path = os.path.join(get_exe_dir(), "config.json")


def load_config():
    try:
        with open(config_path, "r") as f:
            config = json.load(f)
        return config, True
    except Exception as e:
        return e, False


def save_config(config_data):
    with open(config_path, "w", encoding="utf-8") as f:
        f.write(config_data)


def get_version():
    try:
        with open(os.path.join(get_exe_dir(), "version.txt"), "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return ""


def is_debug():
    return os.environ.get("DEBUG") == "1"


def is_admin():
    if platform.system() == "Windows":
        # Windows: 使用 ctypes 检查是否为管理员
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except:
            return False
    else:
        # macOS / Linux: 判断 UID 是否为 0（root）
        return os.geteuid() == 0


def get_update_info():
    try:
        resp = requests.get("https://cnb.cool/feassh/zying-auto/-/git/raw/main/version.json", timeout=10)
        resp.raise_for_status()
        return resp.json(), True
    except Exception as e:
        return e, False


def check_need_update():
    current_version = get_version()

    info, ok = get_update_info()
    if not ok:
        return info, False

    latest_version = info.get("version")
    desc = info.get("desc")

    if latest_version == current_version:
        return "", False

    return desc, True


def kill_process_by_name(process_name):
    try:
        # 遍历所有运行中的进程
        for proc in psutil.process_iter(['name']):
            # 检查进程名是否匹配（不忽略大小写）
            if proc.info['name'] == process_name:
                try:
                    # 强行终止进程
                    proc.kill()
                    return True
                except psutil.AccessDenied:
                    return False
                except psutil.NoSuchProcess:
                    return False
        else:
            return False
    except Exception as e:
        return False


def get_next_page_point(rect):
    pil_image = ImageGrab.grab(bbox=(rect.left - 5, rect.top - 5, rect.right, rect.bottom + 5))
    img_bgr = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
    hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)

    # 灰色方框 HSV 范围
    gh, gs, gv = 107, 23, 181
    lower_gray = np.array([gh - 10, max(gs - 30, 0), max(gv - 30, 0)])
    upper_gray = np.array([gh + 10, min(gs + 30, 255), min(gv + 30, 255)])

    # 蓝色方框 HSV 范围
    bh, bs, bv = 104, 239, 160
    lower_blue = np.array([bh - 10, max(bs - 30, 0), max(bv - 30, 0)])
    upper_blue = np.array([bh + 10, min(bs + 30, 255), min(bv + 30, 255)])

    # 掩膜
    mask_gray = cv2.inRange(hsv, lower_gray, upper_gray)
    mask_blue = cv2.inRange(hsv, lower_blue, upper_blue)

    # 轮廓检测
    contours_gray, _ = cv2.findContours(mask_gray, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours_blue, _ = cv2.findContours(mask_blue, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # 排序：从左到右
    contours_gray = sorted(contours_gray, key=lambda c: cv2.boundingRect(c)[0])
    contours_blue = sorted(contours_blue, key=lambda c: cv2.boundingRect(c)[0])

    current_page = None
    for cnt in contours_blue:
        x, y, w_box, h_box = cv2.boundingRect(cnt)
        if w_box > 22 and h_box > 22:
            current_page = (x, y, w_box, h_box)
            break

    if current_page is None:
        return None

    c_x, _, _, _ = current_page
    next_page = None
    for cnt in contours_gray:
        x, y, w_box, h_box = cv2.boundingRect(cnt)
        if w_box > 22 and h_box > 22:
            if x > c_x:
                next_page = (x, y, w_box, h_box)
                break

    if next_page is None:
        return None

    n_x, n_y, n_w, n_h = next_page
    center_x = n_x + n_w // 2
    center_y = n_y + n_h // 2

    return rect.left + center_x, rect.top + center_y


def save_kw_to_server(kws):
    if kws is None or len(kws) == 0:
        return True

    try:
        data = []
        for kw, img in kws:
            data.append({
                "kw": kw,
                "img": img
            })

        resp = requests.post("https://zying.feassh.workers.dev/insertBatch", json={
            "data": data,
            "token": "feassh-zying-cf-worker-token"
        }, timeout=20)
        resp.raise_for_status()

        return resp.json()["code"] == 0
    except Exception:
        return False
