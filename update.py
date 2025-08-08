import os
import requests
import zipfile
import hashlib
from tqdm import tqdm

import util

UPDATE_ZIP_NAME = "update.zip"  # 下载的文件名


def download_file_with_progress(url, dest_path):
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
    if expected_hash.strip() == "":
        return True

    sha256 = hashlib.sha256()
    try:
        with open(file_path, 'rb') as f:
            for chunk in iter(lambda: f.read(4096), b''):
                sha256.update(chunk)
        file_hash = sha256.hexdigest()
        return file_hash == expected_hash.lower()
    except Exception as e:
        print(f"校验文件失败: {e}")
        return False


def unzip_and_replace(zip_path, extract_to):
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
        return True
    except Exception as e:
        print(f"解压失败: {e}")
        return False


def upgrade():
    current_version = util.get_version()

    info, ok = util.get_update_info()
    if not ok:
        print(info)
        print("更新失败！版本信息获取失败！")
        return

    latest_version = info.get("version")
    update_url = info.get("downloadUrl")
    expected_hash = info.get("hash")

    print(f"当前版本：{current_version}")
    print(f"最新版本：{latest_version}")

    if latest_version == current_version:
        print("已经是最新版本。")
        return

    print("发现新版本，开始下载更新包...")

    if not download_file_with_progress(update_url, UPDATE_ZIP_NAME):
        return

    print("下载完成，正在校验文件完整性...")

    if not verify_file_sha256(UPDATE_ZIP_NAME, expected_hash):
        print("文件校验失败，更新终止。")
        return

    print("校验通过，正在解压更新包...")

    if not unzip_and_replace(UPDATE_ZIP_NAME, util.get_exe_dir()):
        print("解压失败，更新终止。")
        return

    print("更新完成，清理临时文件。")
    os.remove(UPDATE_ZIP_NAME)

    print("更新成功！")


if __name__ == "__main__":
    try:
        upgrade()
        input("")
    except Exception as e:
        print(e)
        input("更新失败！")
