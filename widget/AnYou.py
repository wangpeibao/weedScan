from PyQt5.QtWidgets import QWidget, QGridLayout, QCheckBox


class AnYou(QWidget):
    def __init__(self, anyou):
        super(AnYou, self).__init__()
        self.anyou = anyou
        layout = QGridLayout()
        self.edit_anyou0 = QCheckBox("未在当地烟草批发企业进货")
        self.edit_anyou1 = QCheckBox("销售非法生产的烟草专卖品")
        self.edit_anyou2 = QCheckBox("销售无标志外国卷烟")
        self.edit_anyou3 = QCheckBox("销售专供出口卷烟")
        self.edit_anyou4 = QCheckBox("无烟草专卖品准运证运输烟草专卖品")
        layout.addWidget(self.edit_anyou0, 0, 0)
        layout.addWidget(self.edit_anyou1, 0, 1)
        layout.addWidget(self.edit_anyou2, 0, 2)
        layout.addWidget(self.edit_anyou3, 1, 0)
        layout.addWidget(self.edit_anyou4, 1, 1)

        self.handle_init_info()
        self.setLayout(layout)

    def handle_init_info(self):
        if self.anyou[0]:
            self.edit_anyou0.setChecked(True)
        if self.anyou[1]:
            self.edit_anyou1.setChecked(True)
        if self.anyou[2]:
            self.edit_anyou2.setChecked(True)
        if self.anyou[3]:
            self.edit_anyou3.setChecked(True)
        if self.anyou[4]:
            self.edit_anyou4.setChecked(True)

    def get_anyou_info(self):
        anyou = ["", "", "", "", ""]
        status = False
        if self.edit_anyou0.isChecked():
            status = True
            anyou[0] = self.edit_anyou0.text()
        if self.edit_anyou1.isChecked():
            status = True
            anyou[1] = self.edit_anyou1.text()
        if self.edit_anyou2.isChecked():
            status = True
            anyou[2] = self.edit_anyou2.text()
        if self.edit_anyou3.isChecked():
            status = True
            anyou[3] = self.edit_anyou3.text()
        if self.edit_anyou4.isChecked():
            status = True
            anyou[4] = self.edit_anyou4.text()
        return status, anyou