import ctypes


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


print(f"屏幕缩放比例: {get_scaling_factor():.2f}x")
