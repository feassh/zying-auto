import ctypes
import os
import platform

import sys
import time
from typing import Literal

import psutil
import pyautogui
from rich.console import Console

console = Console()


def get_exe_dir():
    # PyInstaller打包后的程序，sys.executable 是exe路径
    # 脚本运行时，__file__ 是脚本路径
    if getattr(sys, 'frozen', False):
        # 打包后，exe模式
        return os.path.dirname(sys.executable)
    else:
        # 普通脚本模式
        # .. 上级目录
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


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


# print(f"屏幕缩放比例: {get_scaling_factor():.2f}x")
def get_scaling_factor():
    # 设置进程为 DPI 感知
    ctypes.windll.user32.SetProcessDPIAware()

    # 获取主屏幕的 DC（设备上下文）
    dc = ctypes.windll.user32.GetDC(0)

    # 获取水平和垂直 DPI
    dpi_x = ctypes.windll.gdi32.GetDeviceCaps(dc, 88)  # LOGPIXELSX
    dpi_y = ctypes.windll.gdi32.GetDeviceCaps(dc, 90)  # LOGPIXELSY

    # 释放 DC
    ctypes.windll.user32.ReleaseDC(0, dc)

    # 计算缩放比例（默认 DPI 为 96）
    scaling_factor = dpi_x / 96.0
    return scaling_factor


def print_inline(
    text: str,
    color: str = "white",
    style: Literal["bold", "italic", "underline", "none"] = "none",
    newline: bool = True
):
    """彩色文本 + 可覆盖当前行的控制台输出"""
    style_prefix = f"{style} {color}" if style != "none" else color
    end_char = "\n" if newline else "\r"
    console.print(f"[{style_prefix}]{text}[/{style_prefix}]", end=end_char)


def block_input(block: bool):
    """启用或关闭用户输入"""
    ctypes.windll.user32.BlockInput(block)


def safe_click(x, y, duration=0.05):
    try:
        block_input(True)
        pyautogui.moveTo(x, y, duration=duration)
        time.sleep(0.02)
        pyautogui.click()
    finally:
        block_input(False)
        time.sleep(0.05)


def safe_right_click(x, y, duration=0.05):
    try:
        block_input(True)
        pyautogui.moveTo(x, y, duration=duration)
        time.sleep(0.02)
        pyautogui.rightClick()
    finally:
        block_input(False)
        time.sleep(0.05)
