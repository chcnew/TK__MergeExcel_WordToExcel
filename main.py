from merge_excel import merge_excel
from wordtoexcel import wordtoexcel
import tkinter as tk


def set_window():
    # 创建窗口界面
    window = tk.Tk()
    window.title('Excel合并与Word内容导出工具')
    window.attributes("-alpha", 1.0)  # 窗口透明度
    window.resizable(width=False, height=False)  # 窗口是否可以拉长拉宽
    # 获取屏幕尺寸以计算布局参数，使窗口居屏幕中央
    screenwidth = window.winfo_screenwidth()
    screenheight = window.winfo_screenheight()
    width = 658
    height = 404
    align_str = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    window.geometry(align_str)
    # 创建画布
    canvas_window = tk.Canvas(window, width=658, height=504, bd=0, highlightthickness=5)
    canvas_window.pack()
    # 设置标签Label
    lab_sm = tk.Label(window,
                      text='说明：\n|------------------------------------------------\n|1.请选择以工单号命名的文件夹所在的文件夹！ \n|2.word存放形式：*/工单号/处理方案/*.docx   \n|3.excel存放形式：*/工单号/处理方案/*.excel \n|------------------------------------------------\n',
                      fg='black',
                      font=('微软雅黑', 12),
                      width=24,
                      height=10,
                      justify='left')
    canvas_window.create_window(329, 100, width=400, height=150, window=lab_sm)
    # 设置运行按钮
    butt_var = tk.StringVar()
    butt_var.set('MergeExcel')
    butt_ex = tk.Button(window, textvariable=butt_var, font=('微软雅黑', 12), fg='red', command=merge_excel, activebackground='#90EE90', bg='#FFE4B5')
    canvas_window.create_window(240, 240, width=150, height=40, window=butt_ex)
    butt_var1 = tk.StringVar()
    butt_var1.set('WordToExcel')
    butt_doc = tk.Button(window, textvariable=butt_var1, font=('微软雅黑', 12), fg='red', command=wordtoexcel, activebackground='#90EE90', bg='#FFE4B5')
    canvas_window.create_window(429, 240, width=150, height=40, window=butt_doc)
    window.mainloop()


if __name__ == '__main__':
    set_window()
