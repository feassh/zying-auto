import ctypes
import math
import os
import re
import sys
import time
from datetime import date, datetime
from pathlib import Path

import pyautogui
import pyperclip
from pywinauto.application import Application
from rich.console import Console
from typing import Literal
from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

import util

# 设置 Playwright 浏览器路径（相对路径）
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(Path(__file__).parent / "ms-playwright")

console = Console()


def check_load_finished(delay=3.0):
    last_time = time.time()
    while True:
        if (10, 90, 160) == pyautogui.pixel(992, 594):
            last_time = time.time()
        if time.time() - last_time >= delay:
            break


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
        block_input(True)                 # 禁用用户输入
        pyautogui.moveTo(x, y, duration=duration)
        time.sleep(0.02)
        pyautogui.click()
    finally:
        block_input(False)                # 重新启用用户输入
        time.sleep(0.05)


def safe_right_click(x, y, duration=0.05):
    try:
        block_input(True)                 # 禁用用户输入
        pyautogui.moveTo(x, y, duration=duration)
        time.sleep(0.02)
        pyautogui.rightClick()
    finally:
        block_input(False)                # 重新启用用户输入
        time.sleep(0.05)


try:
    # 获取自身控制台窗口句柄
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd:
        screen_width, screen_height = pyautogui.size()
        # SWP_NOZORDER=0x4 保持层次
        ctypes.windll.user32.SetWindowPos(hwnd, 0, 1120, 0, screen_width - 1120, math.ceil(screen_height / 1.2), 0x4)
    else:
        print_inline("找不到自身控制台窗口句柄，请手动将本软件窗口移至没有遮挡目标软件的地方", color="red")

    ################## 读取配置文件 ##################
    config, ok = util.load_config()
    if not ok:
        print_inline("配置文件读取失败！" + config, color="red")
        sys.exit()

    if config['currentPage'] <= 0:
        ################## 启动应用程序 ##################
        print_inline("正在启动目标应用程序...")
        app = Application(backend="uia").start(config['exePath'])

        # 连接到窗口
        login_window = app.window(title="系统登录")

        # 等待窗口出现（重要！）
        login_window.wait("visible", timeout=300)

        # 打印控件树（用于查找控件）
        # login_window.print_control_identifiers()

        print_inline("正在登录...", color="yellow")
        input_username = login_window.child_window(auto_id="txtAcc")
        input_pwd = login_window.child_window(auto_id="txtPwd")
        bt_login = login_window.child_window(auto_id="btnLogin")

        if input_username.window_text() == "":
            input_username.set_focus()  # 确保焦点
            input_username.type_keys(config['user'])  # 输入用户名

        # 更安全的密码输入方式（避免明文记录）
        input_pwd.set_focus()
        input_pwd.type_keys(config['pwd'] + "{ENTER}", with_spaces=True)  # 支持特殊键

        # bt_login.click()

    ################## 进入主界面 ##################
    app = Application(backend="win32").connect(title="分销系统", timeout=10)

    main_window = app.window(title="分销系统")
    main_window.wait("visible", timeout=30)

    print_inline("正在初始化应用程序窗口...", color="green")
    main_window.set_focus()
    main_window.move_window(x=0, y=0, width=1110, height=700, repaint=True)
    time.sleep(1)

    if config['currentPage'] <= 0:
        ################## 操作主界面，进入指定界面 ##################
        print_inline("正在定位到指定页面...", color="blue")
        print_inline("点击【产品】选项卡")
        safe_click(x=296, y=40)
        time.sleep(1)
        print_inline("点击【亚马逊选品】选项卡")
        safe_click(x=56, y=678)
        time.sleep(1)
        print_inline("点击【搜索词排名】选项卡")
        safe_click(x=71, y=618)
        # time.sleep(5)
        check_load_finished()
        print_inline("点击【日本】选项卡")
        safe_click(x=329, y=77)
        # time.sleep(5)
        check_load_finished()
        print_inline("点击【50000+】筛选条件")
        safe_click(x=747, y=159)
        # time.sleep(5)
        check_load_finished()
        print_inline("点击【排名上升+】筛选条件")
        safe_click(x=494, y=265)
        # time.sleep(5)
        check_load_finished()
        print_inline("点击【1万+】筛选条件")
        safe_click(x=655, y=299)
        # time.sleep(5)
        check_load_finished()

    total_item = 0
    total_page = 0
    match = re.search(r'共有\s*(\d+)\s*条记录', main_window.child_window(auto_id="lblTotal").window_text())
    if match:
        total_item = int(match.group(1))
        total_page = math.ceil(total_item / 60)
    if total_item <= 0:
        print_inline("搜索词分页数据获取失败！程序已终止运行。", color="red")
        input("按任意键退出...")
        sys.exit()
    else:
        print_inline(f"共获取到 {total_item} 条搜索词数据，一共 {total_page} 页", color="green")

    ################## 启动浏览器，获取数据 ##################
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False if config['showBrowser'] else True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
            locale="ja-JP",
            viewport={"width": 1280, "height": 720},
        )

        # 隐藏 Playwright 指纹
        context.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', {
          get: () => false,
        });
        """)

        for cur_page in range(0 if config['currentPage'] <= 0 else (config['currentPage'] - 1), total_page):
            # 第一页按钮坐标
            first_x = 155
            first_y = 600

            print_inline(f"正在获取第 {cur_page + 1}/{total_page} 页搜索词列表...")
            main_window.set_focus()

            if cur_page + 1 <= 13:
                safe_click(x=first_x + (cur_page * 31), y=first_y)
            elif cur_page + 1 <= 19:
                safe_click(x=first_x + (13 * 31), y=first_y)
            else:
                safe_click(x=first_x + (13 * 31) + 10, y=first_y)

            # time.sleep(5)
            check_load_finished()
            pyperclip.copy("")
            safe_right_click(x=233, y=464)
            time.sleep(1)
            safe_click(x=287, y=504)
            time.sleep(1)
            kw_list = pyperclip.paste().splitlines()
            print_inline(f"成功获取到 {len(kw_list)} 个搜索词", color="blue")

            saved_kw = []
            i = 0
            for kw in kw_list:
                i += 1
                page = context.new_page()
                print_inline(f"当前正在筛选搜索词 (第 {cur_page + 1}/{total_page} 页, 第 {i}/{len(kw_list)}) 个：" + kw,
                             color="yellow") # , newline=False

                try:
                    page.goto("https://www.amazon.co.jp/s?k=" + kw.strip(), timeout=60000, wait_until='load')
                except Exception as e:
                    print_inline("网络超时，自动筛选下一个搜索词...（若频繁超时，请检查本地网络与VPN）", color="red")
                    page.close()
                    time.sleep(config["fetchDelay"])
                    continue

                try:
                    if "ご迷惑をおかけしています" in page.title():
                        print_inline("遇到亚马逊风控，访问终止，自动筛选下一个搜索词...", color="red")
                        page.close()
                        time.sleep(config["fetchDelay"])
                        continue

                    els = page.query_selector_all("div.udm-primary-delivery-message span.a-text-bold")
                except Exception as e:
                    els = []

                count = 0
                for el in els:
                    match = re.search(r'(\d{1,2})月(\d{1,2})日', el.text_content())
                    if match:
                        month = int(match.group(1))
                        day = int(match.group(2))

                        today = date.today()
                        current_year = today.year

                        try:
                            target_date = date(current_year, month, day)

                            delta = (target_date - today).days
                            if delta >= config['minDateInterval']:
                                count += 1
                        except ValueError:
                            pass
                            # print_inline("非法日期，比如 2月30日 之类", color="red")
                    else:
                        pass
                        # print_inline("未匹配到日期", color="red")

                if count >= config['matchCount']:
                    print_inline("该搜索词符合预期条件，已记录到数据库", color="green") # , newline=False
                    saved_kw.append(kw.strip())
                else:
                    print_inline("该搜索词不符合预期条件，已忽略", color="red") # , newline=False
                page.close()
                time.sleep(config["fetchDelay"])

            # 将该页的数据保存为 Excel
            if len(saved_kw) <= 0:
                print_inline("当前页筛选后的搜索词数量为 0，跳过保存为 Excel 的步骤", color="yellow")
                continue

            # 创建工作簿和工作表
            wb = Workbook()
            ws = wb.active
            ws.title = "亚马逊搜索词"

            # 设置标题
            title = "亚马逊 [日本]区域 搜索词"
            ws['A1'] = title
            ws['A1'].font = Font(bold=True)  # 加粗标题

            # 写入数据和超链接
            for i, text in enumerate(saved_kw, start=2):  # 从第2行开始写
                cell = ws.cell(row=i, column=1, value=text)
                cell.hyperlink = "https://www.amazon.co.jp/s?k=" + text
                cell.style = "Hyperlink"  # 使用超链接样式，默认蓝色带下划线

            # 自动调整列宽（简单版）
            max_len = max(len(text) for text in saved_kw)
            ws.column_dimensions[get_column_letter(1)].width = max_len + 5

            # 保存文件，文件名使用当前日期
            filename = datetime.now().strftime("%Y-%m-%d-%H-%M-%s") + f"-第{cur_page + 1}页" + ".xlsx"
            output_dir = Path(os.path.join(util.get_exe_dir(), "excel"))
            output_dir.mkdir(parents=True, exist_ok=True)
            wb.save(output_dir / filename)
            print_inline(f"当前页筛选后的搜索词数量为 {len(saved_kw)} 个，已将数据存储为 Excel 文件：{filename}", color="green")

        browser.close()
        print_inline("程序执行完毕！", color="green")
        input("按任意键退出...")
except Exception as e:
    print(e)
    print_inline("程序执行出错！", color="red")
    input("按任意键退出...")
