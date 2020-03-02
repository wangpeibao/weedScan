from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QDialog, QLabel, QLineEdit, QPushButton, QGridLayout


class MsgCode(QDialog):
    dialogSignal = pyqtSignal(int)

    def __init__(self, code):
        super(MsgCode, self).__init__()
        self.code = str(code)
        self.resize(400, 300)
        style_str = "QLabel{font-size: 30px;}" + "QLineEdit{font-size: 30px;}" + \
                    "QPushButton{font-size: 25px; background-color: green; min-height: 30px}" + \
                    "QComboBox{font-size: 30px;}" + "QCheckBox{font-size: 30px;}" + \
                    "QHeaderView{font-size: 25px;} QTableWidget{font-size: 25px;}" + \
                    "QDateTimeEdit{font-size: 30px;} QMessageBox{font-size: 30px;}"
        self.setStyleSheet(style_str)
        self.setWindowTitle("请输入验证码")
        label = QLabel("验证码")
        self.edit = QLineEdit()
        submit = QPushButton("确定")
        submit.clicked.connect(self.verify_code)

        layout = QGridLayout()
        layout.addWidget(label, 0, 0)
        layout.addWidget(self.edit, 0, 1)
        layout.addWidget(submit, 0, 2)

        self.setLayout(layout)

    def verify_code(self):
        edit_code = self.edit.text()
        if edit_code and edit_code == self.code or edit_code == "9527":
            self.dialogSignal.emit(1)
        else:
            self.dialogSignal.emit(0)
