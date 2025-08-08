import json
import os
import subprocess
import sys

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox

import util
from ui_main_window import Ui_mainWindow

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

        self.ui.pbExePath.clicked.connect(self.pb_exe_path)
        self.ui.pbStart.clicked.connect(self.pb_start)

        if not util.is_admin():
            QMessageBox.critical(self, "提示", "请使用管理员方式运行本软件")
            QTimer.singleShot(0, QApplication.quit)
            return

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
        })

        util.save_config(config_data)

        self.ui.pbStart.setText("🟢停止自动化🟢")
        self.running = True
        self.process = subprocess.Popen([os.path.join(util.get_exe_dir(), "main.exe")])

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
