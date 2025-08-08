import asyncio
import json
import os
import subprocess
import sys

from PyQt5.QtCore import QTimer, QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox

import util
from ui_main_window import Ui_mainWindow

class AsyncWorker(QThread):
    # 定义信号，用于将异步任务的结果传递回主线程
    result_signal = pyqtSignal(tuple)

    def run(self):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        result = loop.run_until_complete(self.async_task())
        loop.close()

        # 发出信号，将结果传回主线程
        self.result_signal.emit(result)

    async def async_task(self):
        return util.check_need_update()

class MyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_mainWindow()
        self.ui.setupUi(self)

        self.window().setWindowTitle(self.window().windowTitle() + " - " + util.get_version())

        self.process = None
        self.running = False

        config, ok = util.load_config()
        if ok:
            self.ui.leExePath.setText(config["exePath"])
            self.ui.leUser.setText(config["user"])
            self.ui.lePwd.setText(config["pwd"])
            self.ui.sbMinDateInterval.setValue(config["minDateInterval"])
            self.ui.sbMaxDateInterval.setValue(config["maxDateInterval"])
            self.ui.sbMatchCount.setValue(config["matchCount"])
            self.ui.sbFetchDelay.setValue(config["fetchDelay"])
            self.ui.sbConcurrency.setValue(config["concurrency"])
            self.ui.cbShowBrowser.setChecked(config["showBrowser"])
            self.ui.leExcelPath.setText(config["excelPath"])

        self.ui.pbExePath.clicked.connect(self.pb_exe_path)
        self.ui.pbStart.clicked.connect(self.pb_start)
        self.ui.pbExcelPath.clicked.connect(self.pb_excel_path)

        if not util.is_admin():
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
        show_browser = self.ui.cbShowBrowser.isChecked()
        excel_path = self.ui.leExcelPath.text()

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
            "showBrowser": show_browser,
            "excelPath": excel_path
        })

        util.save_config(config_data)

        self.ui.pbStart.setText("🟢停止自动化🟢")
        self.running = True
        self.process = subprocess.Popen([os.path.join(util.get_exe_dir(), "main.exe")])

    def pb_excel_path(self):
        directory_path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if directory_path:
            self.ui.leExcelPath.setText(directory_path)

    def update_app(self, result: tuple):
        desc, ok = result
        if ok:
            msg_box = QMessageBox()
            msg_box.setWindowTitle("提示")
            msg_box.setText("检测到新版本，更新内容：\n\n" + desc)
            msg_box.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
            msg_box.setDefaultButton(QMessageBox.Ok)

            # 设置按钮的文本（可选，中文化）
            msg_box.button(QMessageBox.Ok).setText("更新")
            msg_box.button(QMessageBox.Cancel).setText("取消")

            # 显示消息框并获取用户选择
            result = msg_box.exec_()

            # 根据用户选择处理
            if result == QMessageBox.Ok:
                subprocess.Popen([os.path.join(util.get_exe_dir(), "update.exe")])
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
