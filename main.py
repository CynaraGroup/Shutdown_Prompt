import sys
import os
import platform
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QMessageBox
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont
import win32com.client as win32

class ShutdownNotice(QWidget):
    def __init__(self):
        super().__init__()
        self.remaining = 60  # 先初始化倒计时变量
        self.init_ui()  # 再调用界面初始化

    def init_ui(self):
        # 窗口基本设置
        self.setWindowTitle('关闭提示')
        self.setFixedSize(800, 600)
        # 删除标题栏
        self.setWindowFlags(self.windowFlags() | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)

        # 主布局
        layout = QVBoxLayout()
        layout.setContentsMargins(30, 30, 30, 20)
        layout.setSpacing(20)

        # 标题
        lbl_title = QLabel("离开教室前请关闭")
        lbl_title.setFont(QFont("Microsoft YaHei", 28, QFont.Bold))
        lbl_title.setAlignment(Qt.AlignCenter)

        # 正文内容
        lbl_content = QLabel("空调、灯光、窗户、多媒体\n关机倒计时")
        lbl_content.setFont(QFont("Microsoft YaHei", 24))
        lbl_content.setAlignment(Qt.AlignCenter)

        # 倒计时标签
        self.lbl_countdown = QLabel()
        self.lbl_countdown.setFont(QFont("Microsoft YaHei", 48, QFont.Bold))
        self.lbl_countdown.setAlignment(Qt.AlignCenter)
        self.lbl_countdown.setStyleSheet("color: #ff4444;")

        # 底部信息
        lbl_footer = QLabel("© 2025 Cynara, All rights reserved.\n萌ICP备20250202号")
        lbl_footer.setFont(QFont("Microsoft YaHei", 8))
        lbl_footer.setAlignment(Qt.AlignCenter)
        lbl_footer.setStyleSheet("color: #666666;")

        # 添加组件到布局
        layout.addWidget(lbl_title)
        layout.addWidget(lbl_content)
        layout.addWidget(self.lbl_countdown)
        layout.addStretch()
        layout.addWidget(lbl_footer)
        self.setLayout(layout)

        # 初始化倒计时显示
        self.update_countdown()

        # 定时器设置
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.countdown)
        self.timer.start(1000)  # 每秒更新

    def update_countdown(self):
        self.lbl_countdown.setText(f"{self.remaining} 秒")

    def countdown(self):
        self.remaining -= 1
        if self.remaining >= 0:
            self.update_countdown()
        else:
            self.timer.stop()
            self.save_office_documents()
            self.shutdown_system()

    def save_office_documents(self):
        try:
            # 保存 Powerpoint 文档
            powerpoint = win32.gencache.EnsureDispatch('PowerPoint.Application')
            for presentation in powerpoint.Presentations:
                if presentation.Saved == False:
                    presentation.Save()
            powerpoint.Quit()

            # 保存 Word 文档
            word = win32.gencache.EnsureDispatch('Word.Application')
            for doc in word.Documents:
                if doc.Saved == False:
                    doc.Save()
            word.Quit()
        except Exception as e:
            print(f"保存文档时出错: {e}")

    def shutdown_system(self):
        """执行关机操作"""
        system = platform.system()
        if system == "Windows":
            os.system("shutdown -s -f -t 0")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ShutdownNotice()
    window.show()
    sys.exit(app.exec_())