import cv2
import numpy as np
from PIL import ImageGrab

import config

cv2_window_name = "Image Process Result"


def get_next_page_point(rect):
    global cv2_window_name

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

    if config.DEBUG:
        cv2.drawContours(img_bgr, contours_gray, -1, (30, 96, 139), 2)
        cv2.drawContours(img_bgr, contours_blue, -1, (255, 206, 73), 2)

    current_page = None
    for cnt in contours_blue:
        x, y, w_box, h_box = cv2.boundingRect(cnt)
        if w_box > 22 and h_box > 22:
            current_page = (x, y, w_box, h_box)

            if config.DEBUG:
                cv2.drawContours(img_bgr, [cnt], -1, (255, 0, 0), 3)

            break

    if config.DEBUG:
        cv2.imshow(cv2_window_name, img_bgr)
        cv2.moveWindow(cv2_window_name, 0, 700)
        cv2.waitKey(1)

    if current_page is None:
        return None

    c_x, _, _, _ = current_page
    next_page = None
    for cnt in contours_gray:
        x, y, w_box, h_box = cv2.boundingRect(cnt)
        if w_box > 22 and h_box > 22:
            if x > c_x:
                next_page = (x, y, w_box, h_box)

                if config.DEBUG:
                    cv2.drawContours(img_bgr, [cnt], -1, (0, 255, 0), 3)

                break

    if config.DEBUG:
        cv2.imshow(cv2_window_name, img_bgr)
        cv2.waitKey(1)

    if next_page is None:
        return None

    n_x, n_y, n_w, n_h = next_page
    center_x = n_x + n_w // 2
    center_y = n_y + n_h // 2

    return rect.left + center_x, rect.top + center_y
