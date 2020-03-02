# 扫描录入的窗口

from PyQt5.QtCore import QRegExp, pyqtSignal
from PyQt5.QtGui import QRegExpValidator
from PyQt5.QtWidgets import QApplication, QDialog, QLineEdit, QVBoxLayout, QLabel

class ScanDial(QDialog):
    dialogSignal = pyqtSignal(str)
    def __init__(self):
        super(ScanDial, self).__init__()

        self.resize(200, 100)
        self.setWindowTitle("提示信息")

        self.idLable = QLabel("使用扫码枪或者手动输入并敲入回车")
        self.idEdit = QLineEdit()
        self.idEdit.returnPressed.connect(self.emitYanID)
        idValidator = QRegExpValidator(QRegExp("^[0-9]{13}$"))
        self.idEdit.setValidator(idValidator)

        # 装载布局
        layout = QVBoxLayout()
        layout.addWidget(self.idLable)
        layout.addWidget(self.idEdit)
        self.setLayout(layout)

        style_str = "QLabel{font-size: 30px;}" + "QLineEdit{font-size: 30px;}"
        self.setStyleSheet(style_str)



    # 窗口内出现enter事件
    def emitYanID(self):
        self.dialogSignal.emit(self.idEdit.text())


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    ui = ScanDial()
    ui.show()
    sys.exit(app.exec_())