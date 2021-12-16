import tkinter as tk
from tkinter import filedialog


window = tk.Tk()
window.withdraw()

# 获取文件夹路径
Dir_path = filedialog.askdirectory()
# 获取文件路径
File_path = filedialog.askopenfilename()


print(Fpath)