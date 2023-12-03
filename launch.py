from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QLineEdit, QVBoxLayout, QHBoxLayout, QLabel, QMessageBox
from PyQt6.QtGui import QFont, QPixmap, QPainter, QColor
from PyQt6.QtCore import QTimer
import sys
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import os
from io import BytesIO

class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("科能高级技工学校“双优计划”评分表")
        self.setGeometry(100, 100, 400, 200)

        vLayout = QVBoxLayout()
        hLayout = QHBoxLayout()
        
        #* 选择文件按钮
        select_file_button = QPushButton("选择文件", self)
        select_file_button.setFixedHeight(40)
        select_file_button.clicked.connect(self.select_file_button_click)

        #* 输入框
        self.path_input = QLineEdit(self)
        self.path_input.setPlaceholderText("请点击选择文件或手动键入 Excel 文件")
        font = QFont()
        font.setItalic(True)
        font.setPointSize(12)
        self.path_input.setFont(font)
        self.debounce_timer = QTimer(self)
        self.debounce_timer.setSingleShot(True)
        self.debounce_timer.timeout.connect(self.on_path_input_confirmed)
        self.path_input.textChanged.connect(self.on_path_input_change)
       
        #* 生成按钮
        # todo 消息闪现（成果、失败）
        image_generate_button = QPushButton("生成", self)
        image_generate_button.setFixedHeight(40)
        # file_generate.setStyleSheet("background-color: green")
        image_generate_button.clicked.connect(self.image_generate_button_click)

        #* 导出按钮
        image_export_button = QPushButton("导出", self)
        image_export_button.setFixedHeight(40)
        # file_export.setStyleSheet("background-color: yellow")
        image_export_button.clicked.connect(self.image_export_button_click)

        self.status_label = QLabel(self)
        self.image_label = QLabel(self)  

        vLayout.addWidget(select_file_button)
        vLayout.addWidget(self.path_input)

        hLayout.addWidget(image_generate_button)  
        hLayout.addWidget(image_export_button) 
        vLayout.addLayout(hLayout)

        vLayout.addWidget(self.status_label)
        vLayout.addWidget(self.image_label)  

        self.setLayout(vLayout)

        self.image_generated = False 
        self.file_path = ""

    def select_file_button_click(self):
        file_dialog = QFileDialog()
        file_dialog.setNameFilter("Excel文件 (*.xlsx)")
        self.file_path, _ = file_dialog.getOpenFileName(self, "选择文件", "", "Excel文件 (*.xlsx)")
        if os.path.isfile(self.file_path):
            self.path_input.setText(self.file_path)
        else:
            QMessageBox.warning(self, "警告", "文件路径错误，请选择一个有效的 Excel 文件", QMessageBox.StandardButton.Ok)

    def on_path_input_change(self):
        self.debounce_timer.start(1000)

    def on_path_input_confirmed(self):
        if os.path.isfile(self.path_input.text()):
            self.file_path = self.path_input.text()
        else:
            QMessageBox.warning(self, "警告", "文件路径错误，请选择一个有效的 Excel 文件", QMessageBox.StandardButton.Ok)
        # self.status_label.setText("文本框内容：" + self.path_input.text())
    
    def image_generate_button_click(self):
        # todo: image generate here
        if not self.file_path:
            QMessageBox.warning(self, "警告", "请先「选择文件」", QMessageBox.StandardButton.Ok)
            return

        if self.file_path:
            df = pd.read_excel(self.file_path)
            print("列名:", df.columns)
            # # 提取N12, P12, R12, T12, V12单元格的内容
            # values = [df.at[11, 'N'], df.at[11, 'P'], df.at[11, 'R'], df.at[11, 'T'], df.at[11, 'V']]

            # # 定义饼图的大小，以及每个部分的标签和颜色
            # sizes = [35, 35, 15, 5, 10]
            # labels = ['35%', '35%', '15%', '5%', '10%']
            # colors = ['green', 'red', 'blue', 'yellow', 'orange']
            # explode = (0.1, 0, 0, 0, 0)  # 仅突出显示第一个部分

            # # 计算实际占比
            # actual_sizes = [value / size if value < size else size for value, size in zip(values, sizes)]
            # remaining_sizes = [size - actual for actual, size in zip(actual_sizes, sizes)]

            # # 绘制实际占比的饼图
            # fig, ax = plt.subplots()
            # ax.pie(actual_sizes, explode=explode, labels=labels, colors=colors, autopct='%1.1f%%', startangle=140)

            # # 绘制剩余部分的灰色饼图
            # grey_colors = ['grey' if remaining > 0 else 'white' for remaining in remaining_sizes]
            # ax.pie(remaining_sizes, colors=grey_colors, radius=0.7, startangle=140)
            # ax.set(aspect="equal")

            # 保存图表为QPixmap
            buffer = BytesIO()
            plt.savefig(buffer, format='png')
            buffer.seek(0)
            pixmap = QPixmap()
            pixmap.loadFromData(buffer.getvalue())
            buffer.close()

            # 显示图片
            self.image_label.setPixmap(pixmap)
            self.status_label.setText("已成功生成！请点按「导出」")
            self.image_generated = True
        else:
            QMessageBox.warning(self, "警告", "文件路径为空，请选择一个有效的 Excel 文件", QMessageBox.StandardButton.Ok)

    def image_export_button_click(self):
        if not self.image_generated:  # 如果没有生成图片
            QMessageBox.warning(self, "警告", "请先「生成」图片", QMessageBox.StandardButton.Ok)
            return
        
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getSaveFileName(self, "保存文件", "", "PNG文件 (*.png)")
        if file_path:
            pixmap = self.image_label.pixmap()
            pixmap.save(file_path)
            self.status_label.setText("已导出到：" + file_path)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec())