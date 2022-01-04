import os
from pandas.core import frame
from docx import Document
import time
import datetime
import tkinter as tk
from tkinter import filedialog
import sys


def wordtoexcel():
    # 打开word文档
    # 结合Tk设计
    window = tk.Tk()
    window.withdraw()
    print("请选择需要拆分的Word文件所在的文件夹！")
    path = filedialog.askdirectory()
    names = os.listdir(path)
    lst1 = []
    for i in names:
        if i.split(".")[1] == "docx" or i.split(".")[1] == "doc":
            path1 = path + '/' + i
            document = Document(path1)
            # 获取所有段落
            all_paragraphs = document.paragraphs
            lst = [j.text for j in all_paragraphs]
            lst.insert(0, i)
            lst1.append(lst)
            print('处理完毕：' + i)
            time.sleep(0.0005)

    # 二维列表转Dataframe
    df = frame.DataFrame(lst1)
    df.to_excel('./' + 'WordtoExcel文件' + datetime.datetime.now().strftime('%Y%m%d-%H%M%S') + '.xlsx', index=False)
    sys.exit(0)
