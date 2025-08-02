import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QHBoxLayout, QPushButton, QWidget, QMessageBox
from PyQt5.QtGui import QFont, QPalette, QLinearGradient, QColor, QFontDatabase, QGradient, QPainter, QPainterPath
from PyQt5.QtCore import Qt, QTimer, QRectF

import win32com.client as win32
import os
import platform

class ShutdownPrompt(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_UI()
        self.init_Timer()

    def init_UI(self):
        self.setWindowTitle('离班提示')
        self.setFixedSize(800, 600)

        # 隐藏标题栏、窗口置顶
        self.setWindowFlags(self.windowFlags() | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        # 窗口透明
        self.setAttribute(Qt.WA_TranslucentBackground)

        # 加载外部字体
        font_zqk_in = QFontDatabase.addApplicationFont('assist/font/zhengqingke.ttf')
        font_dt_in = QFontDatabase.addApplicationFont('assist/font/dingtalk.ttf')
        font_zqk = QFontDatabase.applicationFontFamilies(font_zqk_in)
        font_dt = QFontDatabase.applicationFontFamilies(font_dt_in)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        central_widget.setStyleSheet(
            '''
            QWidget {
             border-radius: 30px; 
            }
            '''
        )
        layout = QVBoxLayout(central_widget)
        layout.setAlignment(Qt.AlignCenter)

        # 背景
        gradient = QLinearGradient(0, 0, 1, 1)
        gradient.setCoordinateMode(QGradient.ObjectBoundingMode)
        gradient.setColorAt(0, QColor(163,221,252))
        gradient.setColorAt(1, QColor(244,244,252))
        palette = QPalette()
        palette.setBrush(QPalette.Window, gradient)
        central_widget.setPalette(palette)
        central_widget.setAutoFillBackground(True)
        self.setPalette(palette)
        self.setAutoFillBackground(True)

        # 大标题
        title_label = QLabel('离开教室前请关闭')
        title_font = QFont(font_zqk[0], 48, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setStyleSheet(
            '''
            color: rgb(147, 112, 219);
            '''
        )
        layout.addWidget(title_label, alignment=Qt.AlignCenter)

        # 器材
        equipment_label = QLabel('空调|灯光|窗户|多媒体')
        equipment_font = QFont('SimHei', 36, QFont.Bold)
        equipment_label.setFont(equipment_font)
        equipment_label.setStyleSheet('color: rgb(255, 99, 71);')  # 红色
        layout.addWidget(equipment_label, alignment=Qt.AlignCenter)

        # 创建左侧水平布局
        left_layout = QHBoxLayout()
        left_layout.setSpacing(5)  # 文本与倒计时的间距
        left_layout.setContentsMargins(0, 0, 0, 0)

        # OS即将关闭：
        down_label = QLabel('OS即将关闭：')
        down_font = QFont('SimHei', 24, QFont.Bold)
        down_label.setFont(down_font)
        down_label.setStyleSheet('color: rgb(105, 105, 105); margin: 0px; padding: 0px;')
        down_label.setAlignment(Qt.AlignCenter)  # 设置标签文本居中
        left_layout.addWidget(down_label)

        # 倒计时
        # solve
        countdown_container = QWidget()
        countdown_layout = QVBoxLayout(countdown_container)  # 改为垂直布局
        countdown_layout.setSpacing(20)
        countdown_layout.setContentsMargins(20, 0, 20, 0)
        # 数字渲染
        self.time_label = QLabel('60')
        time_font = QFont(font_dt[0], 64, QFont.Bold)
        self.time_label.setFont(time_font)
        self.time_label.setStyleSheet(
            '''
            color: rgb(255, 105, 180); 
            margin: 0px; padding: 0px;
            '''
        )
        self.time_label.setAlignment(Qt.AlignCenter)  # 设置标签文本居中
        left_layout.addWidget(self.time_label)
        countdown_container = QWidget()
        # 布局
        countdown_layout = QVBoxLayout(countdown_container)
        countdown_layout.setAlignment(Qt.AlignCenter)
        countdown_layout.addLayout(left_layout)

        # 手动关机
        shutdown_btn = QPushButton("手动关机")
        shutdown_btn.setStyleSheet(
            '''
            QPushButton {
                background-color: rgb(180, 180, 180);
                color: white;
                border-radius: 15px;
                padding: 10px 20px;
                font-size: 52px;
                font-family: 'SimHei';
            }
            QPushButton:hover {
                background-color: rgb(160, 160, 160);
            }
            '''
        )
        # 点击事件
        shutdown_btn.clicked.connect(self.show_confirm_dialog)
        countdown_layout.addWidget(shutdown_btn, alignment=Qt.AlignCenter)
        layout.addWidget(countdown_container)

        # 版权信息
        copyright_label = QLabel('© 2025 Cynara, All rights reserved\n萌ICP备20250202号')
        copyright_font = QFont('SimHei', 10)
        copyright_label.setFont(copyright_font)
        copyright_label.setStyleSheet(
            '''
            color: rgb(169, 169, 169);
            '''
        )
        copyright_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(copyright_label)

    # 初始化倒计时
    def init_Timer(self):
        self.remaining_time = 60
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.updateCountdown)
        self.timer.start(1000) # 间隔1s

    # 倒计时更新
    def updateCountdown(self):
        self.remaining_time -= 1
        self.time_label.setText(str(self.remaining_time))
        if self.remaining_time <= 0:
            self.timer.stop()
            self.save_office_documents()
            self.shutdown_system()

    # 手动关机确认
    def show_confirm_dialog(self):
        # 创建自定义提示框
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("提示")
        msg_box.setText("忘记关打死你")
        # 设置自定义按钮文本
        ok_button = msg_box.addButton("行", QMessageBox.AcceptRole)
        msg_box.setDefaultButton(ok_button)
        # 显示对话框并等待用户点击
        msg_box.exec_()
        # 点击按钮后关闭程序
        self.close()

    # 关闭Powerpoint和Word
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

    # 关机
    def shutdown_system(self):
        system = platform.system()
        os.system("shutdown -s -f -t 0")

    # 绘制圆角窗口 - 完善此方法（关键修改）
    def paintEvent(self, event):
        # 创建圆角矩形路径
        radius = 30
        path = QPainterPath()
        rect = QRectF(self.rect())
        path.addRoundedRect(rect, radius, radius)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        # 绘制背景
        gradient = QLinearGradient(0, 0, self.width(), self.height())
        gradient.setColorAt(1, QColor(173, 216, 230))
        gradient.setColorAt(0, QColor(255, 182, 193))
        painter.setPen(Qt.NoPen)
        painter.setBrush(gradient)
        painter.drawPath(path)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ShutdownPrompt()
    window.show()
    sys.exit(app.exec_())