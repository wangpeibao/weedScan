# 显示统计数据并选择不同的打印
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QDialog, QLabel, QPushButton, QGridLayout, QApplication, QCheckBox


class Choose(QDialog):
    dialogSignal = pyqtSignal(int)

    def __init__(self, count_data):

        super(Choose, self).__init__()

        self.resize(250, 200)
        self.setWindowTitle("提示信息")

        # 初始化数据
        self.jiaLableSort = QLabel("假")
        self.jiaLableCount = QLabel("数量: %d" % count_data["jia"]["count"])
        self.jiaLablePrice = QLabel("总价: %.2f" % count_data["jia"]["price"])

        self.siLableSort = QLabel("私")
        self.siLableCount = QLabel("数量: %d" % count_data["si"]["count"])
        self.siLablePrice = QLabel("总价: %.2f" % count_data["si"]["price"])

        self.feiLableSort = QLabel("非")
        self.feiLableCount = QLabel("数量: %d" % count_data["fei"]["count"])
        self.feiLablePrice = QLabel("总价: %.2f" % count_data["fei"]["price"])

        # self.totalPrice = QLabel(str(count_data["price"]))


        # 装载布局
        layout = QGridLayout()
        layout.addWidget(self.jiaLableSort, 0, 0)
        layout.addWidget(self.jiaLableCount, 0, 1)
        layout.addWidget(self.jiaLablePrice, 0, 2)
        layout.addWidget(self.siLableSort, 1, 0)
        layout.addWidget(self.siLableCount, 1, 1)
        layout.addWidget(self.siLablePrice, 1, 2)
        layout.addWidget(self.feiLableSort, 2, 0)
        layout.addWidget(self.feiLableCount, 2, 1)
        layout.addWidget(self.feiLablePrice, 2, 2)
        # layout.addWidget(self.totalPrice, 3, 0, 1, 3)

        self.anjian_type = 0  # 案件的类型 0零盒 1小微 2一般

        # 根据数量的不同，选择类型
        if count_data["is_tiao"]:  # 以条为单位
            if count_data["jia"]["count"] > 0:
                self.anjian_type = 2
            else:
                all_count = count_data["jia"]["count"] + count_data["si"]["count"] + count_data["fei"]["count"]
                price_count = count_data["price"]
                if all_count > 50 or float(price_count) > 10000:
                    self.anjian_type = 2
                else:
                    self.anjian_type = 1
        else:  # 以支为单位
            if count_data["jia"]["count"] > 0:
                self.anjian_type = 2
            else:
                all_count = count_data["jia"]["count"] + count_data["si"]["count"] + count_data["fei"]["count"]
                if all_count > 800:
                    self.anjian_type = 2
                else:
                    self.anjian_type = 0

        self.confirm_btn = QPushButton("确定")
        self.confirm_btn.clicked.connect(self.func_confirm)

        if self.anjian_type == 0:  # 零盒案件
            self.checkbox1 = QCheckBox("当事人无异议")
            self.checkbox2 = QCheckBox("涉案烟未销售")
            layout.addWidget(self.checkbox1, 4, 0, 1, 2)
            layout.addWidget(self.checkbox2, 5, 0, 1, 2)
            layout.addWidget(self.confirm_btn, 6, 0, 1, 2)
        elif self.anjian_type == 1:  # 小微案件
            self.checkbox1 = QCheckBox("当事人无异议")
            layout.addWidget(self.checkbox1, 4, 0, 1, 2)
            layout.addWidget(self.confirm_btn, 5, 0, 1, 2)
        else:  # 一般案件不处理
            layout.addWidget(self.confirm_btn, 4, 0, 1, 2)

        self.setLayout(layout)
        style_str = "QLabel{font-size: 30px;}" + "QLineEdit{font-size: 30px;}" + \
                    "QPushButton{font-size: 25px; background-color: green; min-height: 50px}" + \
                    "QComboBox{font-size: 30px;}" + "QCheckBox{font-size: 30px;}"
        self.setStyleSheet(style_str)

    # 点击确定按钮的操作
    def func_confirm(self):
        if self.anjian_type == 0:  # 零盒
            if self.checkbox1.isChecked() and self.checkbox2.isChecked():
                self.dialogSignal.emit(0)
            else:
                self.dialogSignal.emit(2)
        elif self.anjian_type == 1:  # 小微
            if self.checkbox1.isChecked():
                try:
                    self.dialogSignal.emit(1)
                except Exception as e:
                    print(e)
            else:
                self.dialogSignal.emit(2)
        else:  # 一般
            self.dialogSignal.emit(2)



if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)
    ui = Choose()
    ui.show()
    sys.exit(app.exec_())
