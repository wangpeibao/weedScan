import sqlite3

from PyQt5.QtCore import pyqtSignal, QRegExp
from PyQt5.QtGui import QDoubleValidator, QRegExpValidator
from PyQt5.QtSql import QSqlDatabase, QSqlQuery
from PyQt5.QtWidgets import QApplication, QDialog, QPushButton, QLabel, QGridLayout, QLineEdit, QComboBox, QMessageBox


class ScanHand(QDialog):
    # yanidSignal = pyqtSignal(str, str)
    yaninfoSignal = pyqtSignal(dict)
    def __init__(self, yan_id):
        super(ScanHand, self).__init__()
        self.resize(400, 300)
        self.setWindowTitle("手动输入信息")

        # 连接数据库
        self.db = QSqlDatabase.addDatabase('QSQLITE')
        self.db.setDatabaseName('./app.db')
        self.db.open()
        self.query = QSqlQuery()

        # 定义控件
        self.idLabel = QLabel("条形码")
        self.idEdit = QLineEdit()
        self.idEdit.setText(yan_id if yan_id else "")
        idValidator = QRegExpValidator(QRegExp("^[0-9]{13}$"))
        self.idEdit.setValidator(idValidator)
        self.idEdit.textChanged.connect(self.querySql)
        self.nameLabel = QLabel("名称")
        self.nameEdit = QLineEdit()
        self.priceLabel = QLabel("单价")
        self.priceEdit = QLineEdit()
        pricevalidator = QDoubleValidator()
        pricevalidator.setDecimals(2)
        self.priceEdit.setValidator(pricevalidator)
        self.unitLabel = QLabel("单位")
        self.unitEdit = QComboBox()
        self.unitEdit.addItems(["条", "盒", "支"])
        self.submitBtn = QPushButton("确定")
        self.submitBtn.clicked.connect(self.submitData)

        # 添加到布局
        layout = QGridLayout()
        layout.addWidget(self.idLabel, 0, 0)
        layout.addWidget(self.idEdit, 0, 1)
        layout.addWidget(self.nameLabel, 1, 0)
        layout.addWidget(self.nameEdit, 1, 1)
        layout.addWidget(self.priceLabel, 2, 0)
        layout.addWidget(self.priceEdit, 2, 1)
        layout.addWidget(self.unitLabel, 3, 0)
        layout.addWidget(self.unitEdit, 3, 1)
        layout.addWidget(self.submitBtn, 4, 1)
        self.setLayout(layout)

        # 页面格式
        style_str = "QLabel{font-size: 30px;}" + "QLineEdit{font-size: 30px;}" + \
                    "QPushButton{font-size: 25px; background-color: green; min-height: 30px}" + \
                    "QComboBox{font-size: 30px;}"
        self.setStyleSheet(style_str)

    # 将填充的数据提交
    def submitData(self):
        yan_id = self.idEdit.text()
        yan_name = self.nameEdit.text()
        yan_price = self.priceEdit.text()
        yan_unit = self.unitEdit.currentText()
        if not yan_id or not yan_name or not yan_price or not yan_unit:
            QMessageBox.critical(self, "错误", "请填写全数据", QMessageBox.Yes)
        else:
            self.yaninfoSignal.emit({
                "yan_id": yan_id,
                "yan_name": yan_name,
                "yan_price": yan_price,
                "yan_unit": yan_unit,
                "yan_pinzhong": "卷烟"
            })

    # 首先验证数据库中是否有
    def querySql(self):
        yan_id = self.idEdit.text()
        response = {}
        if len(yan_id) == 13:
            self.query.exec_("select * from yaninfo where yan_id == %s" % yan_id)
            if self.query.next():
                response = {
                    "yan_id": self.query.value(0),
                    "yan_name": self.query.value(1),
                    "yan_price": self.query.value(3),
                    "yan_unit": "条" if "条" in self.query.value(4) else "支",
                    "yan_pinzhong": "卷烟"
                }
            # return response
        if response:
            self.yaninfoSignal.emit(response)

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    ui = ScanHand()
    ui.show()
    sys.exit(app.exec_())