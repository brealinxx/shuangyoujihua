from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QLineEdit, QVBoxLayout, QHBoxLayout, QLabel, QMessageBox
from PyQt6.QtGui import QFont, QPixmap, QPainter, QColor
import sys

class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("科能高级技工学校“双优计划”评分表")
        self.setGeometry(100, 100, 400, 200)

        vLayout = QVBoxLayout()
        hLayout = QHBoxLayout()
        
        select_file_button = QPushButton("选择文件", self)
        select_file_button.setFixedHeight(40)
        select_file_button.clicked.connect(self.select_file_button_click)

        # 输入框
        self.path_input = QLineEdit(self)
        self.path_input.setPlaceholderText("请点击选择文件或手动键入 Excel 文件")
        font = QFont()
        font.setItalic(True)
        font.setPointSize(10)
        self.path_input.setFont(font)
       
        # todo 消息闪现（成果、失败）
        image_generate_button = QPushButton("生成", self)
        image_generate_button.setFixedHeight(40)
        # file_generate.setStyleSheet("background-color: green")
        image_generate_button.clicked.connect(self.image_generate_button_click)

        # todo 
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
        self.image_generated = False  # 用于标记是否已生成图片

        
    def select_file_button_click(self):
        file_dialog = QFileDialog()
        # file_path, _ = file_dialog.getOpenFileName(self, "选择文件")
        # if file_path:
        #     self.path_input.setText(file_path)
        file_dialog.setNameFilter("Excel文件 (*.xlsx)")
        file_path, _ = file_dialog.getOpenFileName(self, "选择文件", "", "Excel文件 (*.xlsx)")
        if file_path:
            self.path_input.setText(file_path)

    def image_generate_button_click(self):
        # todo: image generate here
        pixmap = QPixmap(100, 100)
        pixmap.fill(QColor("black"))

        self.image_label.setPixmap(pixmap)
        self.status_label.setText("已成功生成！请点按「导出」")

        self.image_generated = True

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