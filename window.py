import asyncio
import json
import os
import subprocess
import sys
from typing import Optional, Any

from PyQt5.QtCore import QTimer, QThread, pyqtSignal, QUrl
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox

import util
import config
from ui_main_window import Ui_mainWindow


class AsyncWorker(QThread):
    # 定义信号，用于将异步任务的结果传递回主线程
    result_signal = pyqtSignal(object)

    def run(self):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        result = loop.run_until_complete(self.async_task())
        loop.close()

        # 发出信号，将结果传回主线程
        self.result_signal.emit(result)

    async def async_task(self):
        return util.net.check_need_update()

class MyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_mainWindow()
        self.ui.setupUi(self)

        self.window().setWindowTitle(self.window().windowTitle() + " - " + util.app.get_version())
        self.window().setFixedSize(self.window().size().width(), self.window().size().height())

        self.process = None
        self.running = False

        config_data = config.get_config(throw_exception=False)
        self.ui.leExePath.setText(config_data.get("exePath", ""))
        self.ui.leUser.setText(config_data.get("user", ""))
        self.ui.lePwd.setText(config_data.get("pwd", ""))
        self.ui.sbMinDateInterval.setValue(config_data.get("minDateInterval", 10))
        self.ui.sbMaxDateInterval.setValue(config_data.get("maxDateInterval", 30)) ##########
        self.ui.sbMatchCount.setValue(config_data.get("matchCount", 5))
        self.ui.sbFetchDelay.setValue(config_data.get("fetchDelay", 0)) ###########
        self.ui.sbConcurrency.setValue(config_data.get("concurrency", 5))
        self.ui.sbRetries.setValue(config_data.get("retries", 3))
        self.ui.sbRetryDelay.setValue(config_data.get("retryDelay", 3))
        self.ui.sbTimeout.setValue(config_data.get("timeout", 60))
        self.ui.leExcelPath.setText(config_data.get("excelPath", ""))
        self.ui.cbDebug.setChecked(config_data.get("debug", False))

        self.ui.pbExePath.clicked.connect(self.pb_exe_path)
        self.ui.pbStart.clicked.connect(self.pb_start)
        self.ui.pbExcelPath.clicked.connect(self.pb_excel_path)
        self.ui.pbOpenWebsite.clicked.connect(self.pb_open_website)

        if not util.system.is_admin():
            QMessageBox.critical(self, "提示", "请使用管理员方式运行本软件")
            QTimer.singleShot(0, QApplication.quit)
            return

        # 初始化并启动异步线程
        self.worker = AsyncWorker()
        self.worker.result_signal.connect(self.update_app)  # 连接信号到槽
        self.worker.start()  # 启动线程

    def pb_exe_path(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "可执行文件 (*.exe);;所有文件 (*)")
        if file_path:
            self.ui.leExePath.setText(file_path)

    def pb_start(self):
        if self.running:
            self.ui.pbStart.setText("🟡启动自动化🟡")
            self.running = False

            if self.process is not None:
                try:
                    self.process.kill()
                except Exception as e:
                    print(e)
                finally:
                    self.process = None

            QMessageBox.information(self, "提示", "自动化已停止")
            return

        exe_path = self.ui.leExePath.text()
        user = self.ui.leUser.text()
        pwd = self.ui.lePwd.text()
        min_date_interval = int(self.ui.sbMinDateInterval.value())
        max_date_interval = int(self.ui.sbMaxDateInterval.value())
        match_count = int(self.ui.sbMatchCount.value())
        fetch_delay = int(self.ui.sbFetchDelay.value())
        concurrency = int(self.ui.sbConcurrency.value())
        current_page = int(self.ui.sbCurrentPage.value()) # 该字段只存储，不读取
        retries = int(self.ui.sbRetries.value())
        retry_delay = int(self.ui.sbRetryDelay.value())
        timeout = int(self.ui.sbTimeout.value())
        excel_path = self.ui.leExcelPath.text()
        debug = self.ui.cbDebug.isChecked()

        if not exe_path or not user or not pwd:
            QMessageBox.warning(self, "提示", "请先配置 智赢软件 的相关信息！")
            return

        config_data = json.dumps({
            "exePath": exe_path,
            "user": user,
            "pwd": pwd,
            "minDateInterval": min_date_interval,
            "maxDateInterval": max_date_interval,
            "matchCount": match_count,
            "fetchDelay": fetch_delay,
            "concurrency": concurrency,
            "currentPage": current_page,
            "retries": retries,
            "retryDelay": retry_delay,
            "timeout": timeout,
            "excelPath": excel_path,
            "debug": debug
        })

        e = config.save_config(config_data)
        if isinstance(e, Exception):
            QMessageBox.critical(self, "提示", f"启动失败！配置信息保存失败！\n{e}")
            return

        self.ui.pbStart.setText("🟢停止自动化🟢")
        self.running = True
        self.process = subprocess.Popen([os.path.join(util.system.get_exe_dir(), "main.exe"), "run"])

    def pb_excel_path(self):
        directory_path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if directory_path:
            self.ui.leExcelPath.setText(directory_path)

    def pb_open_website(self):
        QDesktopServices.openUrl(QUrl("https://zying.woc.cool"))

    def update_app(self, result: Optional[tuple[dict[str, Any], str]]):
        if not result:
            return

        new_version, _ = result

        msg_box = QMessageBox()
        msg_box.setWindowTitle("提示")
        msg_box.setText(f'检测到新版本，更新内容：\n\n{new_version.get("desc")}')
        msg_box.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg_box.setDefaultButton(QMessageBox.Ok)

        # 设置按钮的文本（可选，中文化）
        msg_box.button(QMessageBox.Ok).setText("更新")
        msg_box.button(QMessageBox.Cancel).setText("取消")

        # 显示消息框并获取用户选择
        result = msg_box.exec_()

        # 根据用户选择处理
        if result == QMessageBox.Ok:
            subprocess.Popen([os.path.join(util.system.get_exe_dir(), "update.exe"), "run"])
            QTimer.singleShot(0, QApplication.quit)

    def closeEvent(self, event):
        # 窗口关闭前终止子进程
        if self.process is not None:
            try:
                self.process.kill()
            except Exception as e:
                print(e)

        event.accept()

app = QApplication([])
window = MyApp()
window.show()
sys.exit(app.exec_())
