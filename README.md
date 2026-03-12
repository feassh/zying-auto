```bash
pip install pyqt5 pyqt5-tools
pyqt5-tools designer

pyuic5 ui/main.ui -o .\ui_main_window.py
pyuic5 ui/processor.ui -o .\ui_processor_window.py
pyuic5 ui/qr_login.ui -o ui_qr_login_window.py

$Env:DEBUG = 1
```