from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QLineEdit, QVBoxLayout, QHBoxLayout
from PyQt6.QtGui import QFont
import sys

class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("科能高级技工学校“双优计划”评分表")
        self.setGeometry(100, 100, 400, 200)

        vLayout = QVBoxLayout()
        hLayout = QHBoxLayout()
        
        file_button = QPushButton("选择文件", self)
        file_button.clicked.connect(self.on_file_button_click)

        # 输入框
        self.path_input = QLineEdit(self)
        self.path_input.setPlaceholderText("请点击选择文件或手动键入 Excel 文件")
        font = QFont()
        font.setItalic(True)
        font.setPointSize(10)
        self.path_input.setFont(font)
       
        # todo 消息闪现（成果、失败）
        file_generate = QPushButton("生成", self)
        file_generate.setStyleSheet("background-color: green")
        # file_button.clicked.connect(self.on_file_button_click)

        # todo 
        file_export = QPushButton("导出", self)
        file_export.setStyleSheet("background-color: yellow")

        vLayout.addWidget(file_button)
        vLayout.addWidget(self.path_input)

        hLayout.addWidget(file_generate)  
        hLayout.addWidget(file_export) 
        vLayout.addLayout(hLayout)

        self.setLayout(vLayout)

        
    def on_file_button_click(self):
        file_dialog = QFileDialog()
        # file_path, _ = file_dialog.getOpenFileName(self, "选择文件")
        # if file_path:
        #     self.path_input.setText(file_path)
        file_dialog.setNameFilter("Excel文件 (*.xlsx)")
        file_path, _ = file_dialog.getOpenFileName(self, "选择文件", "", "Excel文件 (*.xlsx)")
        if file_path:
            self.path_input.setText(file_path)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec())