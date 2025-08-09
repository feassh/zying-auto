import ctypes
import time

import cv2
import numpy as np
import pywinauto.mouse
from PIL import ImageGrab
from pywinauto import Application


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


# print(f"屏幕缩放比例: {get_scaling_factor():.2f}x")

# app = Application(backend="win32").connect(title="分销系统", timeout=10)
#
# main_window = app.window(title="分销系统")
# main_window.wait("visible", timeout=30)
#
# main_window.set_focus()
# main_window.move_window(x=0, y=0, width=1110, height=700, repaint=True)
#
# rect = main_window.child_window(auto_id="PTurn").rectangle()
# print(rect)
#
# def reco(pil_image):
#     img_bgr = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
#
#     hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
#
#     # bgr_color = np.uint8([[[181, 172, 165]]])  # 注意BGR顺序
#     # hsv_color = cv2.cvtColor(bgr_color, cv2.COLOR_BGR2HSV)[0][0]
#
#     gh, gs, gv = 107, 23, 181
#     # 灰色方框 HSV 范围
#     lower_gray = np.array([gh-10, max(gs-30, 0), max(gv-30, 0)])
#     upper_gray = np.array([gh+10, min(gs+30, 255), min(gv+30, 255)])
#
#     bh, bs, bv = 104, 239, 160
#     lower_blue = np.array([bh-10, max(bs-30, 0), max(bv-30, 0)])
#     upper_blue = np.array([bh+10, min(bs+30, 255), min(bv+30, 255)])
#
#     # 掩膜
#     mask_gray = cv2.inRange(hsv, lower_gray, upper_gray)
#     mask_blue = cv2.inRange(hsv, lower_blue, upper_blue)
#
#     # 轮廓检测
#     contours_gray, _ = cv2.findContours(mask_gray, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
#     cv2.drawContours(img_bgr, contours_gray, -1, (255, 0, 0), 2)
#     contours_blue, _ = cv2.findContours(mask_blue, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
#     cv2.drawContours(img_bgr, contours_blue, -1, (0, 255, 0), 2)
#
#     # 排序：从左到右
#     contours_gray = sorted(contours_gray, key=lambda c: cv2.boundingRect(c)[0])
#     contours_blue = sorted(contours_blue, key=lambda c: cv2.boundingRect(c)[0])
#
#     current_page = None
#     for cnt in contours_blue:
#         x, y, w_box, h_box = cv2.boundingRect(cnt)
#         if w_box > 22 and h_box > 22:
#             current_page = (x, y, w_box, h_box)
#             break
#
#     if current_page is None:
#         print("未识别到当前页")
#         return
#
#     cu_x, cu_y, cu_w, cu_h = current_page
#     next_page = None
#     for cnt in contours_gray:
#         x, y, w_box, h_box = cv2.boundingRect(cnt)
#         if w_box > 22 and h_box > 22:
#             if x > cu_x:
#                 next_page = (x, y, w_box, h_box)
#                 break
#
#     if next_page is None:
#         print("未识别到下一页")
#         return
#
#     ne_x, ne_y, ne_w, ne_h = next_page
#     center_x = ne_x + ne_w // 2
#     center_y = ne_y + ne_h // 2
#     pywinauto.mouse.click(coords=(rect.left + center_x, rect.top + center_y))
#
# for i in range(500):
#     pywinauto.mouse.click(coords=(541, 647))
#     time.sleep(0.5)
#
#     reco(str(i + 1), ImageGrab.grab(bbox=(rect.left - 5, rect.top - 5, rect.right, rect.bottom + 5)))
#     time.sleep(1)

# img_bgr = cv2.cvtColor(np.array(ImageGrab.grab(bbox=(rect.left - 5, rect.top - 5, rect.right, rect.bottom + 5))), cv2.COLOR_RGB2BGR)
# hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
# # 取鼠标点击颜色
# def pick_color(event, x, y, flags, param):
#     if event == cv2.EVENT_LBUTTONDOWN:
#         hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
#         pixel = hsv[y, x]
#         print("HSV:", pixel)
#
# cv2.imshow("hsv", hsv)
# cv2.setMouseCallback("hsv", pick_color)
# cv2.imshow("mask", mask)
