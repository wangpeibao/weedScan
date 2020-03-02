import math
import os

import openpyxl
from PyQt5.QtCore import QUrl
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings
from PyQt5.QtWidgets import QPushButton, QWidget, QGridLayout

from widget.Function import get_yan_info, get_chou_data


class PrintBtn(QPushButton):
    def __init__(self, index, name):
        super(PrintBtn, self).__init__()
        self.index = index
        self.setText(name)


class PrintYiBan(QWidget):
    def __init__(self, base_path):
        super(PrintYiBan, self).__init__()
        self.filename = base_path + "案件信息.xlsx"
        self.base_path = base_path
        wb = openpyxl.load_workbook(self.filename)
        style_str = "QLabel{font-size: 30px;}" + "QLineEdit{font-size: 30px;}" + \
                    "QPushButton{font-size: 25px; background-color: green; min-height: 35px}" + \
                    "QComboBox{font-size: 30px;}" + "QCheckBox{font-size: 30px;}" + \
                    "QHeaderView{font-size: 25px;} QTableWidget{font-size: 25px;}" + \
                    "QDateTimeEdit{font-size: 30px;} QMessageBox{font-size: 30px;}"
        self.setStyleSheet(style_str)

        self.browser = QWebEngineView()
        self.browser.settings().setAttribute(QWebEngineSettings.AllowRunningInsecureContent, True)

        url = "C:/html_tmp/yiban.html"
        self.browser.load(QUrl(url))

        layout = QGridLayout()
        yan_data = get_yan_info(wb)
        row_index = 0
        for i in range(math.ceil(len(yan_data) / 10)):
            print_xbtz = PrintBtn(index=i, name="先保通知单%d" % (i + 1))
            print_xbtz.clicked.connect(self.print_xbtz)
            layout.addWidget(print_xbtz, row_index, 0)
            row_index += 1
        # 现场勘验笔录
        btn_print_ky = QPushButton("现场勘验笔录")
        btn_print_ky.clicked.connect(self.print_ky)
        layout.addWidget(btn_print_ky, row_index, 0)
        row_index += 1
        # 抽样记录
        chou_data = get_chou_data(wb["一般案件"])
        for i in range(math.ceil(len(chou_data) / 15)):
            print_chou = PrintBtn(index=i, name="抽样记录表%d" % (i + 1))
            print_chou.clicked.connect(self.print_chou)
            layout.addWidget(print_chou, row_index, 0)
            row_index += 1

        self.btn_print_fm = QPushButton("打印封面")
        self.btn_print_fm.clicked.connect(self.print_fm)
        layout.addWidget(self.btn_print_fm, row_index + 1, 0)
        self.btn_back = QPushButton("修改数据")
        layout.addWidget(self.btn_back, row_index + 2, 0)
        self.btn_next = QPushButton("下一步")
        layout.addWidget(self.btn_next, row_index + 3, 0)
        layout.addWidget(self.browser, row_index + 4, 0)
        self.setLayout(layout)

    def print_xbtz(self):
        sender = self.sender()
        self.browser.page().runJavaScript("print_xianbaotz%d();" % sender.index)

    def print_chou(self):
        sender = self.sender()
        filename = self.base_path + "一般抽样%d.xlsx" % sender.index
        os.system('ksolaunch %s' % filename)

    def print_fm(self):
        filename = self.base_path + "封面.docx"
        os.system('ksolaunch %s' % filename)

    def print_ky(self):
        filename = self.base_path + "一般勘验.docx"
        os.system('ksolaunch %s' % filename)