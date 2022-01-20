import pandas as pd
import os
import datetime
import time
import tkinter as tk
from tkinter import filedialog
import sys


def merge_excel():
    # 结合Tk设计
    window = tk.Tk()
    window.withdraw()
    print("请选择需要合并的Excel文件所在的文件夹！")
    path = filedialog.askdirectory()
    print("正在处理，请稍后...")
    names = os.listdir(path)
    lst = []
    for i in names:
        if i.split(".")[1] == "xlsx" or i.split(".")[1] == "xls":
            lst.append(path + '/' + i)

    df = pd.read_excel(lst[0], sheet_name=0)
    empty_df = pd.DataFrame(columns=df.columns)

    n = 0
    for x in lst:
        df = pd.read_excel(x, sheet_name=0)
        df.columns = empty_df.columns
        empty_df = empty_df.append(df, ignore_index=True)
        n = n + 1
        time.sleep(0.0005)

    empty_df.to_excel('./' + 'MergeExcel文件' + datetime.datetime.now().strftime('%Y%m%d-%H%M%S') + '.xlsx', index=False)
    print("处理完成。")
    sys.exit(0)
