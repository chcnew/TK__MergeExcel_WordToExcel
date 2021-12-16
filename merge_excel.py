import pandas as pd
import os
import datetime
import time
import tkinter as tk
from tkinter import filedialog
import sys


def merge_excel():
    window = tk.Tk()
    window.withdraw()
    print("请选择工单号命名的文件夹所在的文件夹！")
    print("存放形式：*/工单号/处理方案/*.excel")
    path = filedialog.askdirectory()
    # path = os.getcwd()
    names = os.listdir(path)
    lst = []
    lst1 = []
    lst2 = []
    for i in names:
        if 'WF364363' in i:
            names1 = os.listdir(path + '/' + i + '/处理方案/')
            lstname1 = [(path + '/' + i + '/处理方案/' + p) for p in names1]
            # print(lstname1)
            for j in lstname1:
                if 'xlsx' == j.split(".")[1]:
                    lst1.append(j)
                    lst2.append(i)
    # print(lst1)
    # print(lst2)

    df = pd.read_excel(lst1[0], sheet_name=0)
    df['工单号'] = lst2[0]
    empty_df = pd.DataFrame(columns=df.columns)

    n = 0
    for x, y in zip(lst1, lst2):
        df = pd.read_excel(x, sheet_name=0)
        df['工单号'] = lst2[n]
        empty_df = empty_df.append(df, ignore_index=True)
        n = n + 1
        print("处理完成：" + str(y))
        time.sleep(0.0005)

    empty_df.to_excel('./' + 'MergeExcel文件' + datetime.datetime.now().strftime('%Y%m%d-%H%M%S') + '.xlsx', index=False)
    sys.exit(0)
