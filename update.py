import os
import subprocess
import sys
import tempfile
import time

import requests
import zipfile
import hashlib
from tqdm import tqdm

import util

UPDATE_ZIP_NAME = "update.zip"  # 下载的文件名


def download_file_with_progress(url, dest_path):
    """使用tqdm显示进度的文件下载函数"""
    try:
        with requests.get(url, stream=True) as r:
            r.raise_for_status()
            total_size = int(r.headers.get('content-length', 0))
            with open(dest_path, 'wb') as f, tqdm(
                    total=total_size, unit='B', unit_scale=True, desc='下载进度'
            ) as bar:
                for chunk in r.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        bar.update(len(chunk))
        return True
    except Exception as e:
        print(f"下载文件失败: {e}")
        return False


def verify_file_sha256(file_path, expected_hash):
    """校验文件的SHA256哈希值"""
    if not expected_hash or expected_hash.strip() == "":
        print("警告：未提供哈希值，跳过校验。")
        return True

    sha256 = hashlib.sha256()
    try:
        with open(file_path, 'rb') as f:
            for chunk in iter(lambda: f.read(4096), b''):
                sha256.update(chunk)
        file_hash = sha256.hexdigest()
        print(f"文件哈希: {file_hash}")
        print(f"预期哈希: {expected_hash.lower()}")
        return file_hash == expected_hash.lower()
    except Exception as e:
        print(f"校验文件失败: {e}")
        return False


def apply_update(zip_path, target_dir):
    """
    应用更新的核心函数。
    将路径直接写入批处理文件，避免参数传递问题。
    """
    try:
        # 1. 创建一个临时目录用于解压
        temp_dir = tempfile.mkdtemp()
        print(f"解压到临时目录: {temp_dir}")

        # 2. 解压更新包到临时目录
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # 3. 稍作等待后删除下载的zip包
        time.sleep(1)
        os.remove(zip_path)

        # 4. 定义批处理文件的路径和内容
        bat_path = os.path.join(temp_dir, "update.bat")

        # 批处理内容：将完整路径硬编码进去，更稳定
        # chcp 65001 切换到UTF-8代码页，更好支持中文路径
        # xcopy 的 /I 开关可以在目标不存在时自动创建目录
        # 增加了错误处理和更新后重启应用的功能
        bat_content = f"""@echo off
echo.
echo  ================================================
echo   正在更新程序，请不要关闭此窗口...
echo  ================================================
echo.

REM 等待主程序完全退出
timeout /t 2 /nobreak > nul

REM 从临时目录复制所有文件到目标目录，覆盖现有文件
echo 正在复制新文件...
xcopy "{temp_dir}" "{target_dir}" /E /Y /I /Q
if errorlevel 1 (
    echo.
    echo [错误] 文件复制失败！请尝试手动解压 update.zip 到程序目录。
    pause
    exit
)

REM 删除临时目录
echo 正在清理临时文件...
rmdir /s /q "{temp_dir}"

echo.
echo  ================================================
echo   更新完成！

pause
exit
"""
        # 5. 创建并写入批处理文件，使用 utf-8 编码
        with open(bat_path, "w", encoding="gbk") as bat:
            bat.write(bat_content)

        # 6. 启动批处理并退出当前程序
        # 直接执行bat文件，不带任何参数
        subprocess.Popen(f'"{bat_path}"', shell=True)
        sys.exit(0)  # 正常退出

    except Exception as e:
        print(f"应用更新失败: {e}")


def upgrade():
    """检查并执行升级流程"""
    current_version = util.get_version()

    print("正在检查更新...")
    info, ok = util.get_update_info()
    if not ok:
        print(info)  # 打印错误信息
        print("更新失败！版本信息获取失败！")
        return

    latest_version = info.get("version")
    update_url = info.get("downloadUrl")
    expected_hash = info.get("hash")

    print(f"当前版本: {current_version}")
    print(f"最新版本: {latest_version}")

    if latest_version == current_version:
        print("已经是最新版本，无需更新。")
        return

    print("发现新版本，准备下载更新...")

    if not download_file_with_progress(update_url, UPDATE_ZIP_NAME):
        return

    print("下载完成，正在校验文件完整性...")

    if not verify_file_sha256(UPDATE_ZIP_NAME, expected_hash):
        print("文件校验失败，更新已终止。请删除 update.zip 后重试。")
        try:
            os.remove(UPDATE_ZIP_NAME)
        except OSError:
            pass
        return

    print("校验通过，准备应用更新...")
    apply_update(UPDATE_ZIP_NAME, util.get_exe_dir())


if __name__ == "__main__":
    try:
        upgrade()
    except Exception as e:
        print(f"发生未知错误: {e}")

    # 如果更新流程没有通过 sys.exit() 退出，则会执行到这里
    input("按 Enter 键退出...")
