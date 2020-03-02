# 登录窗口
import os
from datetime import datetime

from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QWidget, QLabel, QGridLayout, QLineEdit, QPushButton, QApplication, QMessageBox
from PyQt5.QtCore import QRect

from widget.Location import Location


class Login(QWidget):
    def __init__(self, width, height):
        super(Login, self).__init__()
        title_label = QLabel("涉烟案件现场执法平台", self)
        user_label = QLabel("用户：", self)
        pass_label = QLabel("密码：", self)
        self.userLineEdit = QLineEdit(self)
        self.passLineEdit = QLineEdit(self)
        # 打卡配置文件，填充数据
        with open("static/tmp.txt") as f:
            line = f.readline()
            self.userLineEdit.setText(line)
            self.passLineEdit.setText("123456")
            f.close()
        self.passLineEdit.setEchoMode(QLineEdit.Password)
        self.btn_login = QPushButton("登录", self)
        # 设置样式
        style_str = "QLabel{font-size: 25px;}" + "QLineEdit{font-size: 25px;}" + \
                    "QPushButton{font-size: 25px; background-color: green}"
        self.setStyleSheet(style_str)
        # 布局类型
        title_label.setGeometry(
            width * 0.37,
            height * 0.3,
            400, 60
        )
        title_label.setStyleSheet("font-size:40px;")
        user_label.setGeometry(
            width * 0.4,
            height * 0.4,
            100, 30
        )
        pass_label.setGeometry(
            width * 0.4,
            height * 0.45,
            100, 30
        )
        self.userLineEdit.setGeometry(
            width * 0.45,
            height * 0.4,
            200, 30
        )
        self.passLineEdit.setGeometry(
            width * 0.45,
            height * 0.45,
            200, 30
        )
        self.btn_login.setGeometry(
            width * 0.4,
            height * 0.5,
            width * 0.05 + 200,
            40
        )

    # login函数
    def handle_login(self):
        username = self.userLineEdit.text()
        password = self.passLineEdit.text()
        if password == "123456":
            # 如果用户登录成功，创建路径，根据日期/用户名创建
            # now_time_str = datetime.now().strftime("%Y%m%d")
            # if now_time_str in os.listdir("record"):
            #     pass
            # else:
            #     os.mkdir("record/" + now_time_str)
            with open("static/tmp.txt", "w") as f:
                f.writelines(username)
                f.close()
            return username
        else:
            QMessageBox.critical(self, "错误", "用户名密码错误", QMessageBox.Yes)
            return None
