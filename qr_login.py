from PyQt5 import QtWidgets, QtGui, QtCore
from ui_qr_login_window import Ui_QRLoginWindow
from qr_login_worker import QRLoginWorker


class QRLoginWindow(QtWidgets.QMainWindow):
    login_success = QtCore.pyqtSignal(dict)

    def __init__(self):
        super().__init__()

        self.ui = Ui_QRLoginWindow()
        self.ui.setupUi(self)

        self.setWindowTitle("微信扫码登录")

        self.worker = None

        self.ui.refresh_btn.clicked.connect(self.start_login)

        self.start_login()

    def start_login(self):
        if self.worker:
            self.worker.stop()

        self.worker = QRLoginWorker()

        self.worker.qr_loaded.connect(self.show_qr)

        self.worker.status_update.connect(self.ui.status_label.setText)

        self.worker.login_success.connect(self.login_ok)

        self.worker.start()

        self.ui.status_label.setText("正在获取二维码...")

    def show_qr(self, img_bytes, ticket):
        pixmap = QtGui.QPixmap()
        pixmap.loadFromData(img_bytes)

        self.ui.qr_label.setPixmap(
            pixmap.scaled(
                220,
                220,
                QtCore.Qt.KeepAspectRatio,
                QtCore.Qt.SmoothTransformation,
            )
        )

    def login_ok(self, data):
        self.login_success.emit(data)

        QtWidgets.QMessageBox.information(
            self,
            "登录成功",
            f"用户: {data.get('username','')}",
        )

        self.close()

    def closeEvent(self, event):
        if self.worker:
            self.worker.stop()

        super().closeEvent(event)
