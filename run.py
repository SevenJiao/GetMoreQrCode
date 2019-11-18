import os
import sys
import tkinter as tk
from tkinter import filedialog

import numpy as np
import pandas as pd
import qrcode

'''
批量生成二维码
'''

qr = qrcode.QRCode(
    error_correction=qrcode.constants.ERROR_CORRECT_H,
    box_size=10,
    border=4
)


class run:
    def __init__(self, window):
        self.window = window
        self.label_path = tk.StringVar()
        self.curr_path = sys.path[0].replace('/', "\\")
        self.label_path.set('当前目录:'+self.curr_path)
        self.sheet_name = tk.StringVar()
        self.data_col = tk.StringVar()
        self.res_dir = tk.StringVar()
        self.label_res = tk.StringVar()
        self.gui()
        self.window.mainloop()

    def choose_file(self):
        file_name = filedialog.askopenfile(title='选择文件',
                                           filetypes=[("excel文件", "*.xlsx"),
                                                      ('excel2003文件', '*.xls')],
                                           initialdir='g:/')
        self.label_path.set('当前选择:'+file_name.name.replace('/', "\\"))

    def check_dir(self, dir):
        pathdir = self.curr_path+'\\'+dir
        isexit = os.path.exists(pathdir)
        if not isexit:
            os.makedirs(pathdir)
        return pathdir

    def insert_start(self, var):
        self.t4.insert(1.0, var)

    def start(self):
        sheet_name = self.sheet_name.get()
        data_col = self.data_col.get()
        res_dir = self.res_dir.get()
        path_dir = self.check_dir(res_dir)
        labelpath = self.label_path.get()[5::]
        df = pd.read_excel(labelpath, sheet_name=sheet_name)
        file = df.loc[:, [data_col]].values.ravel()
        i = 0
        for readline in file:
            qr.add_data(readline)
            qr.make(fit=True)
            img = qr.make_image()
            filename = str(readline)+'.png'
            readline = ''
            img.save(path_dir+'\\'+filename)
            qr.clear()
            i = i+1
            res = '\ndone No. '+str(i)
            self.insert_start(res)
            print('done No. '+str(i))
            self.window.update()
        res_total = 'sucess '+str(i)+'!'
        self.insert_start(res_total)
        print(res_total)
        self.window.update()

    def gui(self):
        # 路径label
        self.l0 = tk.Label(self.window, textvariable=self.label_path,
                           fg='black', height=2, wraplength=400,
                           justify='left', font=('Arial', 12))
        self.l0.place(x=20, y=0, anchor='nw')
        # 选择button
        self.btnChoose = tk.Button(self.window, text='选择excel文件', font=('Arial', 12),
                                   width=10, height=1, command=self.choose_file)
        self.btnChoose.place(x=480, y=5, anchor='nw')
        # edit
        self.l1 = tk.Label(self.window, text='请输入Sheet名:',
                           fg='black', height=2, font=('Arial', 12))
        self.l1.place(x=130, y=50, anchor='nw')
        self.e1 = tk.Entry(self.window, textvariable=self.sheet_name,
                           show=None, font=('Arial', 16))
        self.e1.place(x=260, y=60, anchor='nw')
        self.l2 = tk.Label(self.window, text='请输入数据列名:',
                           fg='black', height=2, font=('Arial', 12))
        self.l2.place(x=130, y=100, anchor='nw')
        self.e2 = tk.Entry(self.window, textvariable=self.data_col,
                           show=None, font=('Arial', 16))
        self.e2.place(x=260, y=110, anchor='nw')
        self.l3 = tk.Label(self.window, text='输出文件夹名称:',
                           fg='black', height=2, font=('Arial', 12))
        self.l3.place(x=130, y=150, anchor='nw')
        self.e3 = tk.Entry(self.window, textvariable=self.res_dir,
                           show=None, font=('Arial', 16))
        self.e3.place(x=260, y=160, anchor='nw')
        # 开始
        self.btnstart = tk.Button(self.window, text='开始生成', font=('Arial', 18),
                                  width=10, height=1, command=self.start)
        self.btnstart.place(x=240, y=200)

        # 结果展示
        self.t4 = tk.Text(self.window, width=600)
        self.t4.place(x=0, y=300)


if __name__ == '__main__':
    w = tk.Tk()
    w.title('批量生成二维码')
    w.geometry('600x600')
    r = run(w)
