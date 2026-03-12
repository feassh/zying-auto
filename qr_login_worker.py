import base64
import requests
from PyQt5 import QtCore

LOGIN_CODE_URL = "https://api.amz123.com/user/v1/account/wechat/login_code"
QR_STATUS_URL = "https://api.amz123.com/user/v1/account/wechat/qrlogin_status"


class QRLoginWorker(QtCore.QThread):

    qr_loaded = QtCore.pyqtSignal(bytes, str)
    status_update = QtCore.pyqtSignal(str)
    login_success = QtCore.pyqtSignal(dict)

    def __init__(self):
        super().__init__()

        self.ticket = None
        self.running = True

    # {'token': 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJhcHBfdWlkIjo1MjM3ODksInV1aWQiOiIxMDM4NTQxNTI0MjU1NjQxNjAwIiwiZXhwIjoxNzc1OTA1MTMwLCJleHBpcmUiOjE3NzU5MDUxMzB9.DUcJaxqsmTZ9p_AqtOskLp3_GKBCsHU4zxOPz5M-Z1E', 'action': 1, 'avatar': 'https://img.amz123.com/upload/avatar/default_avatar.png', 'username': '和蔼的橙子3127', 'expire': 1775905130, 'app_uid': 123, 'role_id_list': [109, 101]}
    def run(self):
        try:
            r = requests.get(LOGIN_CODE_URL, timeout=10)
            data = r.json()["data"]
            self.ticket = data["ticket"]
            img_bytes = base64.b64decode(data["img_data"])
            self.qr_loaded.emit(img_bytes, self.ticket)
        except Exception as e:
            self.status_update.emit(f"获取二维码失败: {e}")
            return

        while self.running:
            try:
                r = requests.post(
                    QR_STATUS_URL,
                    json={"ticket": self.ticket, "type": 3},
                    timeout=10,
                )
                data = r.json()["data"]
                action = data["action"]
                if action == 0:
                    self.status_update.emit("等待扫码")
                elif action == 1:
                    self.login_success.emit(data)
                    break
                elif action == -1:
                    self.status_update.emit("二维码已过期")
                    break
            except Exception as e:
                self.status_update.emit(f"轮询失败: {e}")

            self.sleep(1)

    def stop(self):
        self.running = False