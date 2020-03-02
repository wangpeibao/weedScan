# 上传现场照片
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QImage, QPixmap
from PyQt5.QtWidgets import QWidget, QLabel, QPushButton, QGridLayout
import cv2
import openpyxl
from openpyxl.drawing.image import Image
from PIL import Image as PILImage


class UploadYan(QWidget):
    def __init__(self, filename):
        super(UploadYan, self).__init__()
        self.photo_index = 0  # 用于记录存储了几张图片
        self.powerStatus = False
        self.filename = filename

        # 图像显示
        layout = QGridLayout()
        self.useBtn = QPushButton("上传违规烟草照片")
        self.useBtn.setDisabled(True)
        self.useBtn.clicked.connect(self.useTake)

        self.noBtn = QPushButton("重新拍照")
        self.noBtn.setDisabled(True)
        self.noBtn.clicked.connect(self.restartTake)

        self.camera = QLabel("摄像头准备中。。。。")
        self.startBtn = QPushButton("打开摄像头")
        self.startBtn.clicked.connect(self.startCamera)
        self.takeBtn = QPushButton("拍照")
        self.takeBtn.clicked.connect(self.takePhoto)

        self.nextBtn = QPushButton("结束")

        layout.addWidget(self.camera, 0, 0, 1, 3)

        layout.addWidget(self.startBtn, 1, 0)
        layout.addWidget(self.takeBtn, 1, 1)
        layout.addWidget(self.noBtn, 1, 2)
        layout.addWidget(self.useBtn, 2, 0, 1, 3)
        layout.addWidget(self.nextBtn, 4, 0, 1, 3)

        self.takeBtn.setDisabled(True)
        self.noBtn.setDisabled(True)
        self.useBtn.setDisabled(True)

        self.setLayout(layout)
        self.capture = None
        self.timer = QTimer(self)

        # 绘制文件格式
        style_str = "QLabel{font-size: 30px;}" + "QLineEdit{font-size: 30px;}" + \
                    "QPushButton{font-size: 25px; background-color: green; min-height: 50px}" + \
                    "QComboBox{font-size: 30px;}"
        self.setStyleSheet(style_str)

    def startCamera(self):
        if not self.powerStatus:
            self.takeBtn.setDisabled(False)
            self.noBtn.setDisabled(True)
            self.useBtn.setDisabled(True)
            self.powerStatus = True
            self.startBtn.setText("关闭摄像头")
            self.capture = cv2.VideoCapture(1)
            self.timer.timeout.connect(self.display)
            self.timer.start(100)
        else:
            self.powerStatus = False
            self.startBtn.setText("打开摄像头")
            self.takeBtn.setDisabled(True)
            self.noBtn.setDisabled(True)
            self.useBtn.setDisabled(True)
            self.timer.stop()
            self.capture.release()

    # 释放占用的资源
    def release(self):
        try:
            self.timer.stop()
            self.capture.release()
        except Exception as e:
            print(e)

    # 从摄像头中获取图片信息
    def get_image(self):
        ret, frame = self.capture.read()
        if ret:
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            self.pil_image = PILImage.fromarray(frame)
            height, width = frame.shape[:2]
            img = QImage(frame, width, height, QImage.Format_RGB888)
            return img
        else:
            return None

    # 显示图片
    def display(self):
        img = self.get_image()
        if img:
            img = QPixmap.fromImage(img)
            self.camera.setPixmap(img)
            self.camera.setScaledContents(True)
        else:
            pass


    # 拍照
    def takePhoto(self):
        self.timer.stop()
        self.takeBtn.setDisabled(True)
        self.useBtn.setDisabled(False)
        self.noBtn.setDisabled(False)

    # 重新拍照
    def restartTake(self):
        self.timer.start(100)
        self.takeBtn.setDisabled(False)
        self.useBtn.setDisabled(True)
        self.noBtn.setDisabled(True)

    # 使用该照片并上传
    def useTake(self):
        # 读写excel并存储图片
        self.uploadPhoto()
        self.timer.start(100)
        self.takeBtn.setDisabled(False)
        self.useBtn.setDisabled(True)
        self.noBtn.setDisabled(True)

    # 读写excel数据
    def uploadPhoto(self):
        wb = openpyxl.load_workbook(self.filename)
        try:
            ws = wb["违规烟草照片"]
        except Exception as e:
            print(e)
            ws = wb.create_sheet(title="违规烟草照片")
        self.pil_image.save("./tmp.jpg")
        img_file = Image("./tmp.jpg")
        ws.add_image(img_file, "A" + str(self.photo_index * 30 + 1))
        wb.save(self.filename)
        self.photo_index += 1
