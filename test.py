import pandas as pd
import matplotlib.pyplot
import shutil
import launch 


# print(matplotlib.matplotlib_fname())

df = pd.read_excel("")

pd.set_option('display.max_rows', None)  # 显示所有行
pd.set_option('display.max_columns', None)  # 显示所有列
pd.set_option('display.width', None)  # 自动调整列宽

values = [df.at[38, 'Unnamed: 13'], df.at[38, 'Unnamed: 15'], df.at[38, 'Unnamed: 17'], df.at[38, 'Unnamed: 19'], df.at[38, 'Unnamed: 21']]

print(values)