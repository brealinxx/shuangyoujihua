import pandas as pd
import matplotlib.pyplot
import shutil
import launch 
import openpyxl
import os


# print(matplotlib.matplotlib_fname())

#df = pd.read_excel("")

# pd.set_option('display.max_rows', None)  # 显示所有行
# pd.set_option('display.max_columns', None)  # 显示所有列
# pd.set_option('display.width', None)  # 自动调整列宽
# values = [df.at[38, 'Unnamed: 13'], df.at[38, 'Unnamed: 15'], df.at[38, 'Unnamed: 17'], df.at[38, 'Unnamed: 19'], df.at[38, 'Unnamed: 21']]

workbook = openpyxl.load_workbook('/Users/brealin/test.xlsx',data_only=True)
sheet = workbook.active
# 创建一个空字典来存储每个人的任务和评分
tasks_and_scores = {}
# 遍历每一行，提取姓名、任务和评分
for row in sheet.iter_rows(min_row=4, values_only=True):  # 从第二行开始遍历，假设第一行是标题
    name, task, score = row[5], row[4], row[25]
    if name not in tasks_and_scores:
        tasks_and_scores[name] = {task: score}  # 如果姓名不在字典中，创建一个新的任务和评分字典
    else:
        tasks_and_scores[name][task] = score  # 如果姓名已经在字典中，添加新的任务和评分

# 输出每个人的任务和评分
for name, tasks in tasks_and_scores.items():
    print(f"{name}:")
    for task, score in tasks.items():
        print(f"任务: {task}, 评分: {score}")
print(os.path.exists('/Users/brealin/DEV/shuangyoujihua/background.png'))

#print(GetExcelData("党建项目执行力得分"))