from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QLineEdit, QVBoxLayout, QHBoxLayout, QLabel, QMessageBox, QGridLayout
from PyQt5.QtGui import QFont, QPixmap
from PyQt5.QtCore import QTimer
from PyQt5.QtChart import QChart, QChartView, QPieSeries
import sys
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import openpyxl
import os
from io import BytesIO
import numpy as np

class Window(QWidget):
    def __init__(self):
        super().__init__()
        plt.rcParams['font.sans-serif']=['SimHei'] # move SimHei.ttf file to /path/to/your/virtualenv/lib/pythonX.X/site-packages/matplotlib/mpl-data/fonts/ttf/
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
        def ColorTrans(r,g,b,a):
            return (r/255,g/255,b/255,a)
        
        if not self.file_path:
            QMessageBox.warning(self, "警告", "请先「选择文件」", QMessageBox.StandardButton.Ok)
            return

        if self.file_path:
            df = pd.read_excel(self.file_path)
            colors = [ColorTrans(211,12,18,1.000), ColorTrans(242,92,5,1.000), ColorTrans(242,206,27,1.000), ColorTrans(15,113,242,1), ColorTrans(13,242,5,1.000)]

            # 保存图表为 QPixmap
            buffer = BytesIO()
            fig, axs = plt.subplots(figsize=(16, 9)) 
            axs.axis('off') 
            gs = GridSpec(3, 3)
            axs = [
                [fig.add_subplot(gs[0, 0]), fig.add_subplot(gs[0, 1]), fig.add_subplot(gs[0, 2])],
                [fig.add_subplot(gs[1, 0]), fig.add_subplot(gs[1, 1]), fig.add_subplot(gs[1, 2])],
                [fig.add_subplot(gs[2, 0]), fig.add_subplot(gs[2, 1]), fig.add_subplot(gs[2, 2])]
                ]
            
            def CreatePie(subplot,value,title):
                data = {'部分':['任务完成得分', '目标达成得分', '资金使用得分', '文档规范性得分', '项目执行力得分'],'占比':value}
                df = pd.DataFrame(data)
                patches, texts, autotexts = subplot.pie(df['占比'], labels=df['部分'],autopct=df['部分'], colors=colors) # modify .venv/lib/python3.12/site-packages/matplotlib/axes/_axes.py 3313 line
                plt.setp(texts, color='none')
                subplot.set_title(title, loc='left', color='blue')
                plt.axis('equal')

            values1 = [df.at[9, 'Unnamed: 13'], df.at[9, 'Unnamed: 15'], df.at[9, 'Unnamed: 17'], df.at[9, 'Unnamed: 19'], df.at[9, 'Unnamed: 21']]
            CreatePie(axs[0][0],values1,"党建考核")

            values2 = [df.at[32, 'Unnamed: 13'], df.at[32, 'Unnamed: 15'], df.at[32, 'Unnamed: 17'], df.at[32, 'Unnamed: 19'], df.at[32, 'Unnamed: 21']]
            CreatePie(axs[0][2],values2,"信息化建设考核")

            values3 = [df.at[16, 'Unnamed: 13'], df.at[16, 'Unnamed: 15'], df.at[16, 'Unnamed: 17'], df.at[16, 'Unnamed: 19'], df.at[16, 'Unnamed: 21']]
            CreatePie(axs[1][0],values3,"立德树人考核")

            values4 = [df.at[38, 'Unnamed: 13'], df.at[38, 'Unnamed: 15'], df.at[38, 'Unnamed: 17'], df.at[38, 'Unnamed: 19'], df.at[38, 'Unnamed: 21']]
            CreatePie(axs[1][2],values4,"社会服务能力考核")

            categories = ['任务完成率', '目标达成度', '资源使用率', '文档规范性', '项目执行力']
            values5 = [20, 30, 25, 15, 10]  #! 对应数据
            axs[1][1].set_aspect('equal')
            # ax_middle.set_xlim(0, 60)  # 设置x轴范围
            axs[1][1].set_ylim(-1.8, len(categories) * 3.5) 
            axs[1][1].set_yticks(np.arange(len(categories)) * 3.3)  # 设置y轴刻度
            axs[1][1].set_yticklabels(categories)  
            bar_height = 3  # 以加粗条形、间距
            for i, (value, category) in enumerate(zip(values5, categories)):
                axs[1][1].barh(i * 3, value, height=bar_height, color=colors[i], edgecolor='none')
                axs[1][1].barh(i * 3, 100-value, left=value, height=bar_height, color='white', edgecolor='none')
            axs[1][1].set_title('整体考核',loc='left', color='blue') 
            # ax_middle.tick_params(axis='y', which='both', left=False)  # 设置y轴刻度参数

            axs[2][0].set_aspect('equal')
            axs[2][0].set_ylim(-1.8, len(categories) * 3.5) 
            axs[2][0].set_yticks(np.arange(len(categories)) * 3.3)  # 设置y轴刻度
            axs[2][0].set_yticklabels(categories)  
            for i, (value, category) in enumerate(zip(values5, categories)):
                axs[2][0].barh(i * 3.2, value, height=bar_height, color=colors[i], edgecolor='none')
                axs[2][0].barh(i * 3.2, 100-value, left=value, height=bar_height, color='white', edgecolor='none')
            axs[2][0].set_title('治理体系考核',loc='left', color='blue') 

            axs[0][1].axis('off') #! test delete frame

            gs.update(wspace=.5, hspace=.3) # 调整整体间距
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