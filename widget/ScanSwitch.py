from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QApplication, QDialog, QPushButton, QLabel, QGridLayout


class ScanSwitch(QDialog):
    dialogSignal = pyqtSignal(int, str)
    def __init__(self, yan_id):
        super(ScanSwitch, self).__init__()
        self.yan_id = yan_id

        self.resize(200, 100)
        self.setWindowTitle("提示信息！")

        btn_scan = QPushButton("继续扫描")
        btn_scan.clicked.connect(self.chooseScan)
        btn_hand = QPushButton("手动输入")
        btn_hand.clicked.connect(self.chooseHand)
        label = QLabel("未查询到相关数据")
        layout = QGridLayout()
        layout.addWidget(label, 0, 0, 1, 0)
        layout.addWidget(btn_scan, 1, 0)
        layout.addWidget(btn_hand, 1, 1)
        self.setLayout(layout)

        # 页面格式
        style_str = "QLabel{font-size: 30px;}" + "QLineEdit{font-size: 30px;}" + \
                    "QPushButton{font-size: 25px; background-color: green; min-height: 30px}" + \
                    "QComboBox{font-size: 30px;}"
        self.setStyleSheet(style_str)

    def chooseScan(self):
        self.dialogSignal.emit(0, self.yan_id)

    def chooseHand(self):
        self.dialogSignal.emit(1, self.yan_id)

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    ui = ScanSwitch()
    ui.show()
    sys.exit(app.exec_())