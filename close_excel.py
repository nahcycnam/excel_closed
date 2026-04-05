import os
import sys
from PyQt6.QtWidgets import QApplication, QMessageBox
from PyQt6.QtGui import QIcon, QPixmap, QColor
from PyQt6.QtCore import Qt


def show_fix_message():
    app = QApplication(sys.argv)

    # 创建消息框
    msg = QMessageBox()
    msg.setWindowTitle('操作提示')

    # 主文本和详细文本
    msg.setText('<b>问题已修复</b>')
    msg.setInformativeText('<span style="font-size: 13px;">请重新打开Excel</span>')

    # 高级样式美化
    msg.setStyleSheet("""
        QMessageBox {
            background-color: white;
            border-radius: 8px;
            border: 1px solid #d0d7e3;
            font-family: 'Segoe UI', 'Microsoft YaHei';
            min-width: 350px;
        }
        QMessageBox QLabel#qt_msgbox_label {
            color: #2c3e50;
            font-size: 16px;
            margin-bottom: 8px;
        }
        QMessageBox QLabel#qt_msgbox_informativelabel {
            color: #7f8c8d;
            font-size: 13px;
            margin-top: 4px;
        }
        QMessageBox QPushButton {
            background-color: #3498db;
            color: white;
            border: none;
            padding: 8px 24px;
            border-radius: 4px;
            min-width: 100px;
            font-size: 14px;
        }
        QMessageBox QPushButton:hover {
            background-color: #2980b9;
        }
        QMessageBox QPushButton:pressed {
            background-color: #1c6ca8;
            padding-top: 9px;
            padding-bottom: 7px;
        }
    """)

    # 添加自定义按钮
    ok_button = msg.addButton("好的", QMessageBox.ButtonRole.AcceptRole)
    ok_button.setCursor(Qt.CursorShape.PointingHandCursor)

    # 设置窗口标志 - 隐藏标题栏但保留窗口阴影
    msg.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)

    # 启用窗口背景（覆盖透明设置）
    msg.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)

    # 显示消息框并处理按钮点击
    msg.exec()

    # 点击"好的"按钮后直接退出应用
    sys.exit()


if __name__ == '__main__':
    os.system('taskkill /im EXCEL.EXE /F')
    show_fix_message()
