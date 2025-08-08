一个 **“可随身携带的离线包”** —— 意思是打包后的 `.exe` 程序：

* ✅ 不依赖目标机器上是否安装 Python
* ✅ 不依赖目标机器是否联网
* ✅ ✅ 自带 Playwright 的 Chromium 浏览器
* ✅ 一次性打包为可执行文件或文件夹，随时拷贝使用

---

## ✅ 最终目标

**打包后 `.exe` 可离线运行、带浏览器、无需额外下载或安装。**

---

## 🧰 工具准备

### 你需要：

1. Python 环境（开发用）

2. 安装好依赖：

   ```bash
   pip install playwright pyinstaller
   playwright install chromium
   ```

3. 找到 Chromium 浏览器的路径（Playwright 下载的）

---

## 🛠️ 实现步骤详解

---

### ✅ 第一步：准备你的 Python 脚本

比如 `main.py`：

```python
import os
from pathlib import Path
from playwright.sync_api import sync_playwright

# 设置 Playwright 浏览器路径（相对路径）
current_dir = Path(__file__).parent
browser_path = current_dir / "browsers"

os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browser_path)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()
    page.goto("https://example.com")
    print("打开网页成功")
    input("按任意键退出...")
```

---

### ✅ 第二步：拷贝 Playwright 浏览器文件夹

Playwright 会把浏览器下载到一个缓存目录，如：

* Windows: `C:\Users\<你>\AppData\Local\ms-playwright\`
* macOS/Linux: `~/.cache/ms-playwright/`

#### 拷贝方式：

将整个 `ms-playwright` 文件夹中的内容（如 `chromium-*`）复制到你的项目文件夹内，例如放在：

```
myproject/
├── main.py
├── browsers/
│   └── chromium-<version>/
```

---

### ✅ 第三步：打包 PyInstaller 命令（重点！）

```bash
pyinstaller --onefile \
    --add-data "browsers;browsers" \
    --collect-all playwright \
    main.py
```

说明：

* `--add-data "browsers;browsers"`：把本地的 `browsers` 文件夹打包进去（注意 Windows 用分号 `;`，Linux/mac 用冒号 `:`）
* `--collect-all playwright`：让 PyInstaller 收集 Playwright 的内部依赖（比如 .node、.wasm 文件）
* `--onefile`：可选。会打包成一个单独的 `.exe`。如果启动速度太慢你可以不用它。

---

### ✅ 第四步：运行测试

打包完成后：

* 查看 `dist/main.exe`
* 双击运行，无需联网，应该可以打开 Chromium 并访问网页
* 检查是否有错误，比如找不到浏览器或启动失败

---

## 📦 可选：打成文件夹而不是 `.exe`

如果你觉得启动慢（因为 `--onefile` 会解压文件），可以用这种方式：

```bash
pyinstaller main.py \
    --add-data "browsers;browsers" \
    --collect-all playwright
```

然后你会得到：

```
dist/
└── main/
    ├── main.exe
    ├── browsers/
    └── ...
```

把整个 `main/` 文件夹复制到 U 盘即可。

---

## 📁 文件结构建议（打包前）：

```
your_project/
├── main.py
├── browsers/               ← 包含 chromium-* 文件夹
│   └── chromium-XXXXX/
├── icon.ico (可选)
```

---

## ✅ 总结：离线携带打包方案关键点

| 要素        | 说明                      |
| --------- | ----------------------- |
| 打包命令      | 使用 `--add-data` 加入浏览器目录 |
| 浏览器路径     | 放在程序目录内，环境变量指向那里        |
| 网络依赖      | **不需要联网**，浏览器文件已内置      |
| Python 环境 | **不需要安装**，已包含解释器        |
| 可执行文件     | `.exe` 可单文件或文件夹形式携带     |

---

## 最终使用的命令

```bash
pip install pyqt5 pyqt5-tools
pyqt5-tools designer
pyuic5 ui/main.ui -o .\ui_main_window.py
```