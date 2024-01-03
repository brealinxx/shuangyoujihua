import pandas as pd
import matplotlib.pyplot
import shutil
import launch 
import openpyxl
import os


# print(matplotlib.matplotlib_fname())

df = pd.read_excel("/Users/brealin/（测试）双优计划任务分解表12.13(1).xlsx")

# pd.set_option('display.max_rows', None)  # 显示所有行
# pd.set_option('display.max_columns', None)  # 显示所有列
# pd.set_option('display.width', None)  # 自动调整列宽
values = [df.at[138, 'Unnamed: 7'], df.at[38, 'Unnamed: 15'], df.at[38, 'Unnamed: 17'], df.at[38, 'Unnamed: 19'], df.at[38, 'Unnamed: 21']]

workbook = openpyxl.load_workbook('/Users/brealin/（测试）双优计划任务分解表12.13(1).xlsx',data_only=True)
sheet = workbook['sheet']
cell_value = sheet.cell(row=140, column=8).value

print(cell_value)

