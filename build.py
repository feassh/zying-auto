import os
import sys
import shutil
import subprocess
from datetime import datetime

# --- 配置 ---
DIST_DIR = "dist"
BUILD_DIR = "build"
SPEC_DIR = "spec"

# 定义脚本的绝对路径
MAIN_SCRIPT = os.path.abspath("main.py")
MAIN_NAME = "主程序"  # 窗口化可执行文件的名称
UPDATE_SCRIPT = os.path.abspath("update.py")


def run_subprocess(command):
    """运行一个子进程命令并实时打印其输出。"""
    try:
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',  # 指定编码以防止错误
            errors='replace',
            bufsize=1
        )
        for line in process.stdout:
            print(line, end="")
        process.wait()
        return process.returncode
    except FileNotFoundError:
        print(f"错误：找不到命令 - {command[0]}。PyInstaller 是否已安装并添加到您的 PATH 环境变量中？")
        return -1
    except Exception as e:
        print(f"运行子进程时发生错误: {e}")
        return -1


def clean():
    """删除之前的构建产物和缓存文件。"""
    print("\n=== 清理旧文件 ===")

    for path in [DIST_DIR, BUILD_DIR, SPEC_DIR]:
        if os.path.exists(path):
            try:
                if os.path.isfile(path):
                    os.remove(path)
                else:
                    shutil.rmtree(path)
                print(f"已删除: {path}")
            except OSError as e:
                print(f"删除 {path} 时出错: {e}")
                raise e

    print("清理完成。")


def ensure_dirs():
    """确保构建和分发目录存在。"""
    os.makedirs(DIST_DIR, exist_ok=True)
    os.makedirs(BUILD_DIR, exist_ok=True)
    os.makedirs(SPEC_DIR, exist_ok=True)


def generate_spec():
    """
    为 PyInstaller 生成一个 .spec 文件，该文件支持多个 EXE 共享一组通用依赖项。
    """
    print("\n=== 生成 .spec 文件 ===")

    spec_content = f"""# -*- mode: python ; coding: utf-8 -*-
# 此文件由构建脚本自动生成。请勿手动编辑。

from PyInstaller.building.build_main import COLLECT, EXE, PYZ, Analysis

block_cipher = None

# --- main.py (窗口应用) 的分析 ---
# 所有数据和隐藏的导入都直接传递给构造函数。
a1 = Analysis(
    [{repr(MAIN_SCRIPT)}],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)
pyz1 = PYZ(a1.pure, a1.zipped_data, cipher=block_cipher)
exe1 = EXE(
    pyz1,
    a1.scripts,
    [], # 二进制文件由 COLLECT 处理
    [], # PYZ 文件由 COLLECT 处理
    [], # 数据文件由 COLLECT 处理
    name='{MAIN_NAME}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False, # 这是一个窗口应用程序
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    uac_admin=True,  # 告诉 Windows 此程序需要管理员权限
)

# --- update.py 的分析 ---
a3 = Analysis(
    [{repr(UPDATE_SCRIPT)}],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)
pyz3 = PYZ(a3.pure, a3.zipped_data, cipher=block_cipher)
exe3 = EXE(
    pyz3,
    a3.scripts,
    [],
    [],
    [],
    name='update',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# --- 将所有输出收集到单个目录中 ---
# 这会创建一个 'shared' 文件夹，其中包含多个 EXE 和所有通用依赖项，
# 从而显著减小总体积。
coll = COLLECT(
    exe1,
    a1.binaries,
    a1.zipfiles,
    a1.datas,
    exe3,
    a3.binaries,
    a3.zipfiles,
    a3.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='shared'
)
"""
    spec_path = os.path.join(SPEC_DIR, "multi.spec")
    with open(spec_path, "w", encoding="utf-8") as f:
        f.write(spec_content)
    print(f"成功生成: {spec_path}")
    return spec_path


def write_version(dist_path):
    """根据当前时间戳生成 version.txt 文件。"""
    print("\n=== 写入版本文件 ===")
    version = datetime.now().strftime("%Y%m%d.%H.%M")
    version_file_path = os.path.join(dist_path, "version.txt")

    # 先写入根目录，然后复制
    with open("version.txt", "w", encoding="utf-8") as f:
        f.write(version)

    shutil.copy("version.txt", version_file_path)
    print(f"版本 '{version}' 已写入 {version_file_path}")
    return version


def compress_output(source_dir, version):
    """将输出目录压缩成一个 zip 文件。"""
    print("\n=== 压缩输出目录 ===")
    # 确保用于存放 zip 文件的基础 dist 目录存在
    os.makedirs(DIST_DIR, exist_ok=True)
    zip_path_base = os.path.join(DIST_DIR, f"zying-auto-v{version}")

    try:
        shutil.make_archive(zip_path_base, 'zip', root_dir=source_dir)
        print(f"成功创建压缩包: {zip_path_base}.zip")
    except Exception as e:
        print(f"创建 zip 压缩包时出错: {e}")


def main():
    """主构建流程执行。"""
    clean()
    ensure_dirs()

    spec_file = generate_spec()

    print("\n=== 开始 PyInstaller 构建过程 ===")
    print("这可能需要几分钟时间...")

    # 注意：--distpath 是将包含 'shared' 文件夹的目录。
    pyinstaller_command = [
        "pyinstaller",
        spec_file,
        "--noconfirm",
        "--distpath", DIST_DIR,
        "--workpath", BUILD_DIR
    ]

    return_code = run_subprocess(pyinstaller_command)

    if return_code != 0:
        print("\n--- 构建失败！ ---")
        print(f"PyInstaller 以错误代码退出: {return_code}")
        sys.exit(return_code)

    print("\n--- 构建成功！ ---")

    dist_shared_path = os.path.join(DIST_DIR, "shared")
    if os.path.exists(dist_shared_path):
        version = write_version(dist_shared_path)
        compress_output(dist_shared_path, version)
        print("\n所有任务均已成功完成。")
    else:
        print(f"\n错误：构建后未找到输出目录 '{dist_shared_path}'。")
        sys.exit(1)


if __name__ == "__main__":
    main()
