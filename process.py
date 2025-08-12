from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QTextCursor, QTextCharFormat, QColor
from PyQt5.QtWidgets import QMainWindow

from processor.search import SearchProcessor
from ui_processor_window import Ui_processorWindow


class AsyncWorker(QThread):
    # 定义信号，用于将异步任务的结果传递回主线程
    log_signal = pyqtSignal(str, str)
    progress_signal = pyqtSignal(int)

    def __init__(self):
        super().__init__()
        # 初始化时添加一个停止标志
        self._should_stop = False

    # 提供一个方法来从外部请求停止
    def request_stop(self):
        self._should_stop = True

    # 提供一个方法来检查是否应该停止
    def is_stopping(self) -> bool:
        return self._should_stop

    def run(self):
        SearchProcessor(self).start_work()


class ProcessWindow(QMainWindow):
    process_window_closed = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.ui = Ui_processorWindow()
        self.ui.setupUi(self)

        self.setFixedSize(self.size().width(), self.size().height())

        # 添加一个标志，防止重复处理关闭事件
        self._is_closing = False

        self.worker = AsyncWorker()
        self.worker.log_signal.connect(self.append_colored_text)
        self.worker.progress_signal.connect(self.update_progress)
        # 将线程的 finished 信号连接到窗口的 close 槽，这样当线程完成后，会自动尝试关闭窗口
        # self.worker.finished.connect(self.close)
        self.worker.start()

    def update_progress(self, value):
        self.ui.progressBar.setValue(value)

    def append_colored_text(self, text, color):
        # 移动光标到文本末尾
        cursor = self.ui.teLog.textCursor()
        cursor.movePosition(QTextCursor.End)

        # 设置文本格式（颜色）
        fmt = QTextCharFormat()
        fmt.setForeground(QColor(color))

        # 插入带颜色的文本
        cursor.insertText(text, fmt)

        # 滚动到光标处（末尾）
        self.ui.teLog.setTextCursor(cursor)
        self.ui.teLog.ensureCursorVisible()

    # 重写 closeEvent 实现优雅退出
    def closeEvent(self, event):
        # 检查线程是否仍在运行，并且还没有开始关闭流程
        if self.worker and self.worker.isRunning():
            if not self._is_closing:
                # 1. 标记正在关闭
                self._is_closing = True

                # 2. 请求后台线程停止
                self.worker.request_stop()

                # 3. 在 UI 上通知用户
                self.append_colored_text("\n\n关闭请求已发送，请等待当前任务完成...\n", "orange")
                self.setWindowTitle("正在退出...")

            # 4. 忽略当前的关闭事件，等待线程自己结束
            event.ignore()
        else:
            # 如果线程已停止，或者已在关闭流程中，则接受事件，正常关闭
            event.accept()

            self.worker = None
            self.process_window_closed.emit()

    def move_to_bottom_right(self):
        """该方法只能在 Window show 之后调用，不然 self.frameGeometry().height() 获取到的值依然是不包括标题栏的"""
        # 不包含任务栏的区域
        available_geometry = self.screen().availableGeometry()

        screen_width = available_geometry.width()
        screen_height = available_geometry.height()

        # 获取窗口自身大小
        window_width = self.width()
        window_height = self.frameGeometry().height()

        # 计算右下角位置
        x = screen_width - window_width
        y = screen_height - window_height

        # 移动窗口
        self.move(x, y)

    def __del__(self):
        print("__del__")
