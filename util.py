import ctypes
import json
import os
import platform
import sys

import requests


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
