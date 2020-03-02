import os
from datetime import datetime
from random import randint

from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QWidget, QDesktopWidget, QGridLayout, QMainWindow, QPushButton, QMessageBox

from widget.Choose import Choose
from widget.DataShow import DataShow
from widget.Function import send_sms
from widget.LingHeKY import LingHeKY
from widget.Location import Location
from widget.Login import Login
import sys

from widget.MsgCode import MsgCode
from widget.PrintLingHe import PrintLingHe
from widget.PrintXiaoWei import PrintXiaoWei
from widget.PrintYiBan import PrintYiBan
from widget.ScanInfo import ScanWidget
from widget.UploadCard import UploadCard
from widget.UploadScene import UploadScene
from widget.UploadYan import UploadYan
from widget.XiaoWeiXB import XiaoWeiXB
from widget.YiBanCY import YiBanCY


class Main(QMainWindow):
    def __init__(self):
        super(Main, self).__init__()
        style_str = "QMessageBox{font-size: 30px;} QPushButton{font-size: 25px}"
        self.setStyleSheet(style_str)
        desktop = QDesktopWidget()
        screen_width = desktop.screenGeometry().width() * 1
        screen_height = desktop.screenGeometry().height() * 0.9
        self.setFixedSize(screen_width, screen_height)
        # 基础地址
        self.base_path = "C:/html_tmp/"
        # self.base_path = "/home/wang/html_tmp/"
        # 需要显示的页面
        self.widget_login = None
        self.widget_location = None
        self.widget_card = None
        self.widget_scene = None
        self.widget_scan = None
        self.widget_choose = None
        self.widget_lingheky = None
        self.widget_xiaowei = None
        self.widget_yiban = None
        self.widget_print_linghe = None
        self.widget_print_xiaowei = None
        self.widget_print_yiban = None
        self.widget_data = None
        self.widget_msgcode = None
        self.widget_uploadyan = None
        self.show_login()
        # self.base_path = "C:/html_tmp/20191114/李根/天津市西青区李七庄街道武台馨苑2227/"
        # self.show_data()

    def show_login(self):
        self.widget_login = Login(self.width(), self.height())
        self.widget_login.btn_login.clicked.connect(self.show_location)
        self.setCentralWidget(self.widget_login)

    def show_location(self):
        username = self.widget_login.handle_login()
        if username:
            # 创建文件目录
            try:
                self.base_path += datetime.now().strftime("%Y%m%d") + "/" + username + "/"
                os.makedirs(self.base_path)
            except Exception as e:
                print(e)
            self.widget_location = Location(base_path=self.base_path, username=username)
            self.widget_location.btn_submit.clicked.connect(self.show_card)
            self.setCentralWidget(self.widget_location)

    def show_card(self):
        self.base_path = self.widget_location.handle_info()
        self.widget_card = UploadCard(self.base_path + "案件信息.xlsx")
        self.widget_card.nextBtn.clicked.connect(self.show_scene)
        self.setCentralWidget(self.widget_card)

    def show_scene(self):
        can_next = self.widget_card.release()
        if not can_next:
            return
        self.widget_scene = UploadScene(self.base_path + "案件信息.xlsx")
        self.widget_scene.nextBtn.clicked.connect(self.show_scan)
        self.setCentralWidget(self.widget_scene)

    def show_scan(self):
        is_legal = self.widget_scene.release()
        if is_legal:
            self.show_login()
        else:
            self.widget_scan = ScanWidget(self.base_path + "案件信息.xlsx")
            self.widget_scan.finishBtn.clicked.connect(self.show_data)
            self.setCentralWidget(self.widget_scan)


    def show_data(self):
        self.widget_scan.uploadData()
        self.widget_data = DataShow(self.base_path + "案件信息.xlsx")
        self.widget_data.confirm_btn.clicked.connect(self.show_choose)
        self.setCentralWidget(self.widget_data)

    def show_choose(self):
        QMessageBox.about(self, "提示", "一经确认，无法更改数据")
        yisong_msg = self.widget_data.get_yisong_type()
        if yisong_msg:
            QMessageBox.about(self, "移送提示", yisong_msg)
        self.widget_choose = Choose(self.widget_data.count_data)
        self.widget_choose.dialogSignal.connect(self.handle_choose)
        self.widget_choose.show()

    def handle_choose(self, stype):
        self.widget_choose.close()
        print("案件类型", stype)
        if stype == 0:  # 零盒案件
            self.widget_lingheky = LingHeKY(self.base_path)
            self.widget_lingheky.btn_finish.clicked.connect(self.show_print_linghe)
            self.setCentralWidget(self.widget_lingheky)
        elif stype == 1:  # 小微案件
            self.widget_xiaowei = XiaoWeiXB(self.base_path)
            self.widget_xiaowei.btn_finish.clicked.connect(self.show_print_xiaowei)
            self.setCentralWidget(self.widget_xiaowei)
        else:  # 一般案件
            self.widget_yiban = YiBanCY(self.base_path)
            self.widget_yiban.btn_finish.clicked.connect(self.show_print_yiban)
            self.setCentralWidget(self.widget_yiban)

    def show_print_linghe(self):
        finish_status = self.widget_lingheky.handle_info()
        if finish_status:
            self.widget_print_linghe = PrintLingHe(self.base_path)
            self.widget_print_linghe.btn_back.clicked.connect(self.show_msgcode)
            self.widget_print_linghe.btn_next.clicked.connect(self.show_uploadyan)
            self.setCentralWidget(self.widget_print_linghe)

    def show_print_xiaowei(self):
        finish_status = self.widget_xiaowei.handle_info()
        if finish_status:
            self.widget_print_xiaowei = PrintXiaoWei(self.base_path)
            self.widget_print_xiaowei.btn_back.clicked.connect(self.show_msgcode)
            self.widget_print_xiaowei.btn_next.clicked.connect(self.show_uploadyan)
            self.setCentralWidget(self.widget_print_xiaowei)

    def show_print_yiban(self):
        finish_status = self.widget_yiban.handle_info()
        if finish_status:
            self.widget_print_yiban = PrintYiBan(self.base_path)
            self.widget_print_yiban.btn_back.clicked.connect(self.show_msgcode)
            self.widget_print_yiban.btn_next.clicked.connect(self.show_uploadyan)
            self.setCentralWidget(self.widget_print_yiban)

    # 展示验证码
    def show_msgcode(self):
        a = QMessageBox.question(None, "确认修改", "确认修改数据将会给管理员发送验证码", QMessageBox.Yes | QMessageBox.No)
        if a == QMessageBox.Yes:
            code = randint(1000, 9999)
            send_sms(code)
            self.widget_msgcode = MsgCode(code)
            self.widget_msgcode.dialogSignal.connect(self.handle_msgcode)
            self.widget_msgcode.show()

    # 处理验证码的结果
    def handle_msgcode(self, status):
        self.widget_msgcode.close()
        if status:
            self.widget_scan = ScanWidget(self.base_path + "案件信息.xlsx")
            self.widget_scan.finishBtn.clicked.connect(self.show_data)
            self.setCentralWidget(self.widget_scan)

    # 展示上传违规烟草照片
    def show_uploadyan(self):
        self.widget_uploadyan = UploadYan(self.base_path + "案件信息.xlsx")
        self.widget_uploadyan.nextBtn.clicked.connect(self.handle_finish)
        self.setCentralWidget(self.widget_uploadyan)

    def handle_finish(self):
        self.widget_uploadyan.release()
        self.show_login()

def main():
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("static/img/logo.ico"))
    ui = Main()
    ui.setWindowTitle("涉烟案件现场执法平台")
    ui.show()
    exit_code = app.exec_()
    if exit_code == 888:
        ui.close()
        main()
    else:
        sys.exit()


if __name__ == "__main__":
    main()
