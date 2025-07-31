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
        self.initUI()
        self.initTimer()

    def initUI(self):
        # 设置窗口标题和大小
        self.setWindowTitle('离开教室提示')
        self.setFixedSize(800, 500)

        # 删除标题栏
        self.setWindowFlags(self.windowFlags() | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        # 添加窗口透明属性（关键修改）
        self.setAttribute(Qt.WA_TranslucentBackground)

        # 添加资源路径处理函数
        def resource_path(relative_path):
            """获取资源的绝对路径（支持开发和打包后两种模式）"""
            if hasattr(sys, '_MEIPASS'):
                return os.path.join(sys._MEIPASS, relative_path)
            return os.path.join(os.path.abspath('.'), relative_path)

        # 加载自定义字体
        font_zhengqingke = QFontDatabase.addApplicationFont(resource_path('zhengqingke.ttf'))
        font_dingtalk = QFontDatabase.addApplicationFont(resource_path('dingtalk.ttf'))
        font_families1 = QFontDatabase.applicationFontFamilies(font_zhengqingke) if font_zhengqingke != -1 else ['SimHei']
        font_families2 = QFontDatabase.applicationFontFamilies(font_dingtalk) if font_dingtalk != -1 else ['SimHei']

        # 创建中心部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 确保中心部件内容不会溢出圆角
        central_widget.setStyleSheet('QWidget { border-radius: 30px; }')
        layout = QVBoxLayout(central_widget)
        layout.setAlignment(Qt.AlignCenter)

        # 设置渐变背景到中心部件
        gradient = QLinearGradient(0, 0, 1, 1)  # 相对坐标：右上角(1,0)到右下角(1,1)
        gradient.setCoordinateMode(QGradient.ObjectBoundingMode)  # 使用相对坐标模式
        gradient.setColorAt(1, QColor(173, 216, 230))  # 起始颜色 淡蓝
        gradient.setColorAt(0, QColor(255, 182, 193))  # 结束颜色 淡粉
        palette = QPalette()
        palette.setBrush(QPalette.Window, gradient)
        central_widget.setPalette(palette)
        central_widget.setAutoFillBackground(True)  # 确保背景被填充
        self.setPalette(palette)
        self.setAutoFillBackground(True)  # 确保主窗口背景也被填充

        # 创建标题标签
        title_label = QLabel('离开教室前请关闭')
        title_font = QFont(font_families1[0], 48, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setStyleSheet('color: rgb(147, 112, 219);')  # 紫色
        layout.addWidget(title_label, alignment=Qt.AlignCenter)

        # 创建提示项标签
        items_label = QLabel('空调|灯光|窗户|多媒体')
        items_font = QFont('SimHei', 36, QFont.Bold)
        items_label.setFont(items_font)
        items_label.setStyleSheet('color: rgb(255, 99, 71);')  # 红色
        layout.addWidget(items_label, alignment=Qt.AlignCenter)

        # 创建倒计时标签容器
        countdown_container = QWidget()
        countdown_layout = QVBoxLayout(countdown_container)  # 改为垂直布局
        countdown_layout.setSpacing(20)
        countdown_layout.setContentsMargins(20, 0, 20, 0)

        # 创建左侧水平布局（用于OS标签和倒计时）
        left_layout = QHBoxLayout()
        left_layout.setSpacing(5)  # 保持文本与倒计时的紧凑间距
        left_layout.setContentsMargins(0, 0, 0, 0)

        # 创建"OS即将关闭："标签
        os_label = QLabel('OS即将关闭：')
        os_font = QFont('SimHei', 24, QFont.Bold)
        os_label.setFont(os_font)
        os_label.setStyleSheet('color: rgb(105, 105, 105); margin: 0px; padding: 0px;')
        os_label.setAlignment(Qt.AlignCenter)  # 设置标签文本居中
        left_layout.addWidget(os_label)

        # 创建倒计时数字标签
        self.time_label = QLabel('60')
        time_font = QFont(font_families2[0], 64, QFont.Bold)
        self.time_label.setFont(time_font)
        self.time_label.setStyleSheet('color: rgb(255, 105, 180); margin: 0px; padding: 0px;')
        self.time_label.setAlignment(Qt.AlignCenter)  # 设置标签文本居中
        left_layout.addWidget(self.time_label)

        # 创建倒计时区域容器
        countdown_container = QWidget()
        
        # 创建倒计时区域主布局（垂直）
        countdown_layout = QVBoxLayout(countdown_container)  # 直接将布局绑定到容器
        countdown_layout.setAlignment(Qt.AlignCenter)  # 设置布局内控件居中
        
        # 添加水平布局（文本和倒计时）
        countdown_layout.addLayout(left_layout)  # 将left_layout添加到垂直布局
        
        # 添加手动关机按钮
        shutdown_btn = QPushButton("手动关机")
        shutdown_btn.setStyleSheet('''
            QPushButton {
                background-color: rgb(180, 180, 180);
                color: white;
                border-radius: 15px;
                padding: 10px 20px;
                font-size: 52px;
                font-family: 'SimHei'; /* 设置黑体 */
            }
            QPushButton:hover {
                background-color: rgb(160, 160, 160);
            }
        ''')
        # 修改点击事件连接
        shutdown_btn.clicked.connect(self.show_confirm_dialog)
        countdown_layout.addWidget(shutdown_btn, alignment=Qt.AlignCenter)  # 添加按钮到垂直布局

        layout.addWidget(countdown_container)

        # 创建版权信息标签
        copyright_label = QLabel('© 2025 Cynara, All rights reserved\n萌ICP备20250202号')
        copyright_font = QFont('SimHei', 10)
        copyright_label.setFont(copyright_font)
        copyright_label.setStyleSheet('color: rgb(169, 169, 169);')  # 浅灰色
        copyright_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(copyright_label)

    def initTimer(self):
        self.remaining_time = 60
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.updateCountdown)
        self.timer.start(1000)  # 每秒更新一次

    def updateCountdown(self):
        self.remaining_time -= 1
        self.time_label.setText(str(self.remaining_time))
        if self.remaining_time <= 0:
            self.timer.stop()
            self.save_office_documents()
            self.shutdown_system()

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