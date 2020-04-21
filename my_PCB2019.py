import pandas as pd
import numpy as np
from pandas import Series, DataFrame

#pd.options.display.max_columns = 999

# 指定文件
file_name_2019PCB = 'd:\PCB\软通MAG项目过程能力基线2019年度PCB - V0.21 (20200407)(开发).xlsm'
excel_file = pd.ExcelFile(file_name_2019PCB)
# 读取文件的所有表单名，得到列表
sheet_names = excel_file.sheet_names

df_2019_development = excel_file.parse(sheet_name='2019开发', header=[0,1], index_col=[0,1], )
df_2019_test = excel_file.parse(sheet_name='2019测试', header=[0,1], index_col=[0,1])
df_2019_autotest = excel_file.parse(sheet_name='2019自动化', header=[0,1], index_col=[0,1])
df_2019_maintenance = excel_file.parse(sheet_name='2019维护', header=[0,1], index_col=[0,1])

df_2019_development.index.names = ['m_index_class', 'm_index_index']
df_2019_development.columns.names = ['m_columns_BU', 'm_columns_PCB']
df_2019_development = df_2019_development.iloc[:, 0:5]

for i in excel_file.sheet_names[9:]:
    df_data = excel_file.parse(sheet_name=i, header=[0,1], index_col=[0,1])
    df_data.index.names = ['m_index_class', 'm_index_index']
    df_data.columns.names = ['m_columns_BU', 'm_columns_PCB']
#    df_data.columns = df_data.columns.map(lambda x: str(x)+i)
    df_2019_development = pd.merge(df_2019_development, df_data.iloc[:, 10:15], \
        how='left', left_index=True, right_index=True, sort=False, suffixes=('_BG', ('_'+i)))
    print(i)
print(df_2019_development.columns)
""" print('--------------------------merge------------------------')
print(df_2019_development.head(3))
 """
df_2019_development.to_excel('D:\PCB\df_2019_development.xlsx', sheet_name='df_2019_development')


