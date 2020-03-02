import openpyxl
from PyQt5.QtWidgets import QWidget, QTableWidget, QHeaderView, QTableWidgetItem, QPushButton, QLabel, QGridLayout

from widget.Function import get_yan_info


class DataShow(QWidget):
    def __init__(self, filename):
        super(DataShow, self).__init__()
        self.filename = filename
        # 绘制页面格式
        style_str = "QLabel{font-size: 30px;}" + "QLineEdit{font-size: 30px;}" + \
                    "QPushButton{font-size: 25px; background-color: green; min-height: 30px}" + \
                    "QComboBox{font-size: 30px;}" + \
                    "QHeaderView{font-size: 25px;} QTableWidget{font-size: 25px;}" + \
                    "QMessageBox{font-size: 30px;}"
        self.setStyleSheet(style_str)
        # 表格空间
        self.table_info = QTableWidget()
        self.table_info.setColumnCount(7)
        self.table_info.setHorizontalHeaderLabels(["条码", "名称", "单价", "单位", "分类", "数量", "总价"])
        self.table_info.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 获取统计数据
        self.count_data = self.handel_data()
        # 确认数据按钮和统计数据显示
        self.confirm_btn = QPushButton("确认")
        self.title_count = QLabel("数量统计：%d" % self.count_data["count"])
        self.jia_count = QLabel("假：%d" % self.count_data["jia"]["count"])
        self.si_count = QLabel("私：%d" % self.count_data["si"]["count"])
        self.fei_count = QLabel("非：%d" % self.count_data["fei"]["count"])

        self.title_price = QLabel("总价统计：%d" % self.count_data["price"])
        self.jia_price = QLabel("假：%d" % self.count_data["jia"]["price"])
        self.si_price = QLabel("私：%d" % self.count_data["si"]["price"])
        self.fei_price = QLabel("非：%d" % self.count_data["fei"]["price"])

        layout = QGridLayout()
        layout.addWidget(self.table_info, 0, 0, 1, 4)
        layout.addWidget(self.title_count, 1, 0)
        layout.addWidget(self.jia_count, 1, 1)
        layout.addWidget(self.si_count, 1, 2)
        layout.addWidget(self.fei_count, 1, 3)
        layout.addWidget(self.title_price, 2, 0)
        layout.addWidget(self.jia_price, 2, 1)
        layout.addWidget(self.si_price, 2, 2)
        layout.addWidget(self.fei_price, 2, 3)
        layout.addWidget(self.confirm_btn, 3, 0, 1, 4)

        self.setLayout(layout)

    # 展示汇总的数据
    def handel_data(self):
        wb = openpyxl.load_workbook(self.filename)
        results = get_yan_info(wb)
        count_data = {
            "jia": {"count": 0, "price": 0},
            "si": {"count": 0, "price": 0},
            "fei": {"count": 0, "price": 0},
            "count": 0,
            "price": 0,
            "is_tiao": False,
            "has_card": True
        }
        self.table_info.setRowCount(len(results))
        for index, result in enumerate(results):
            # 渲染数据
            self.table_info.setItem(index, 0, QTableWidgetItem(result["yan_id"]))
            self.table_info.setItem(index, 1, QTableWidgetItem(result["yan_name"]))
            self.table_info.setItem(index, 2, QTableWidgetItem(result["yan_price"]))
            self.table_info.setItem(index, 3, QTableWidgetItem(result["yan_unit"]))
            self.table_info.setItem(index, 4, QTableWidgetItem(result["yan_sort"]))
            self.table_info.setItem(index, 5, QTableWidgetItem(result["yan_count"]))
            self.table_info.setItem(index, 6, QTableWidgetItem(result["yan_total"]))
            # 统计数据
            count_data["count"] += int(result["yan_count"])
            count_data["price"] += float(result["yan_total"])
            if result["yan_sort"] == "假":
                count_data["jia"]["count"] += int(result["yan_count"])
                count_data["jia"]["price"] += float(result["yan_total"])
            elif result["yan_sort"] == "非":
                count_data["fei"]["count"] += int(result["yan_count"])
                count_data["fei"]["price"] += float(result["yan_total"])
            else:
                count_data["si"]["count"] += int(result["yan_count"])
                count_data["si"]["price"] += float(result["yan_total"])
            if "条" in result["yan_unit"]:
                count_data["is_tiao"] = True
        # 判断有证与否
        ws = wb["许可证照片"]
        if ws["A1"].value == "无证":
            count_data["has_card"] = False
        return count_data

    def get_yisong_type(self):
        if self.count_data["jia"]["price"] >= 150000:
            return "本次执法检查假冒卷烟案值达到15万以上，涉嫌构成销售伪劣产品罪，应当向公安机关移送此案件"
        if self.count_data["price"] >= 50000:
            if not self.count_data["has_card"]:
                return "本次执法检查当事人无烟草专卖零售许可证，同时违法卷烟案值达到5万元以上，涉嫌构成非法经营罪，应当向公安机关移送此案件"
            else:
                return "本次执法检查，违法卷烟案值达到五万元以上，请及时联系辖区公安部门共同研判是否构成涉烟刑事犯罪"
        return ""




