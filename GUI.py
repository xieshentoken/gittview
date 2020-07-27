import re
import sys
from collections import OrderedDict
from itertools import permutations
from tkinter import *
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import xlrd

from process import Gitt


class App():
    def __init__(self, master):
        self.master = master
        self.initWidgets()

        self.example = 0
        self.tao = 3600
        self.massload = 1
        self.actarea = 1
        self.density = 1
        self.DROP = 0
        self.customize_Constant = 2   #  待定
        self.result = {'discharge': pd.DataFrame(), 'charge': pd.DataFrame()}

        self.rgb = ('#000000', 'black')
        
    def initWidgets(self):
        # 初始化菜单、工具条用到的图标
        self.init_icons()
        # 调用init_menu初始化菜单
        self.init_menu()
    #----------------------------------------------------------------------------------
    # 创建第一个容器
        fm1 = Frame(self.master)
        # 该容器放在左边排列
        fm1.pack(side=TOP, fill=BOTH, expand=NO)

        ttk.Label(fm1, text='File Path:', font=('StSong', 20, 'bold')).pack(side=LEFT, ipadx=5, ipady=5, padx= 10)
        # 创建字符串变量，用于传递文件地址
        self.excel_adr = StringVar()
        # 创建Entry组件，将其textvariable绑定到self.excel_adr变量
        ttk.Entry(fm1, textvariable=self.excel_adr,
            width=24,
            font=('StSong', 20, 'bold'),
            foreground='#8080c0').pack(side=LEFT, ipadx=5, ipady=5)#fill=BOTH, expand=YES)
        ttk.Button(fm1, text='...',
            command=self.open_filename # 绑定open_filename方法
            ).pack(side=LEFT, ipadx=1, ipady=5)
    #-----------------------------------------------------------------------------------
    # 创建第二个容器
        fm2 = Frame(self.master)
        fm2.pack(side=TOP, fill=BOTH, expand=YES)
        self.DROPcheck= StringVar()  #  定义变量 用.get()获取
        DROP_check = Scale(fm2, label='前向舍点', from_=0, to=1, orient=HORIZONTAL,
            length=60, showvalue=100, tickinterval=0.5, resolution=1,
            variable = self.DROPcheck,
            command=None)
        DROP_check.pack(side=LEFT, ipadx=1, ipady=5, padx=20, pady=10)
        work_button = Button(fm2, text = 'Work', 
            bd=3, width = 10, height = 1, 
            command = self.work, 
            activebackground='black', activeforeground='white')
        work_button.pack(side=LEFT, ipadx=1, ipady=5, padx=5, pady=10)
        ud_button = Button(fm2, text = 'U-DLi+', 
            bd=3, width = 10, height = 1, 
            command = self.UD_plot, 
            activebackground='black', activeforeground='white')
        ud_button.pack(side=LEFT, ipadx=1, ipady=5, padx=55, pady=10)
        new_button = Button(fm2, text = 'New Path', 
            bd=3, width = 10, height = 1, 
            command = self.new_path, 
            activebackground='black', activeforeground='white')
        new_button.pack(side=LEFT, ipadx=1, ipady=5, padx=15, pady=10)
    #---------------------------------------------------------------------------------------  
        
    # 创建menubar
    def init_menu(self):
        # '初始化菜单的方法'
        # 定义菜单条所包含的3个菜单
        menus = ('文件', '编辑', '参数', '帮助')
        # 定义菜单数据
        items = (OrderedDict([
                # 每项对应一个菜单项，后面元组第一个元素是菜单图标，
                # 第二个元素是菜单对应的事件处理函数
                ('新建', (None, self.new_project)),
                ('打开', (None, self.open_filename)),
                ('另存为', OrderedDict([('CSV', (None, None)),
                        ('Excel',(None, self.saveToexcel))])),
                ('-1', (None, None)),
                ('退出', (None, self.master.quit)),
                ]),
            OrderedDict([('预览',(None, self.preview)), 
                ('-1',(None, None)),
                ('U-DLi+',(None, self.UD_plot)),
                ('Q-DLi+ ',(None, None)),
                ('-2',(None, None)),
                # 二级菜单
                ('更多', OrderedDict([
                    ('选择颜色',(None, self.select_color))
                    ]))
                ]),
            OrderedDict([('弛豫时间τ',(None,self.tao_set)),
                ('-1',(None, None)),
                ('活性物质载量m',(None,self.massLoad_set)),
                ('电化学活性面积A',(None,self.actArea_set)),
                ('活性物质密度ρ',(None,self.density_set))
                ]),
            OrderedDict([('帮助主题',(None, self.original_data_preparation)),
                ('-1',(None, None)),
                ('关于', (None, self.show_help))]))
        # 使用Menu创建菜单条
        menubar = Menu(self.master)
        # 为窗口配置菜单条，也就是添加菜单条
        self.master['menu'] = menubar
        # 遍历menus元组
        for i, m_title in enumerate(menus):
            # 创建菜单
            m = Menu(menubar, tearoff=0)
            # 添加菜单
            menubar.add_cascade(label=m_title, menu=m)
            # 将当前正在处理的菜单数据赋值给tm
            tm = items[i]
            # 遍历OrderedDict,默认只遍历它的key
            for label in tm:
                # print(label)
                # 如果value又是OrderedDict，说明是二级菜单
                if isinstance(tm[label], OrderedDict):
                    # 创建子菜单、并添加子菜单
                    sm = Menu(m, tearoff=0)
                    m.add_cascade(label=label, menu=sm)
                    sub_dict = tm[label]
                    # 再次遍历子菜单对应的OrderedDict，默认只遍历它的key
                    for sub_label in sub_dict:
                        if sub_label.startswith('-'):
                            # 添加分隔条
                            sm.add_separator()
                        else:
                            # 添加菜单项
                            sm.add_command(label=sub_label,image=None,
                                command=sub_dict[sub_label][1], compound=LEFT)
                elif label.startswith('-'):
                    # 添加分隔条
                    m.add_separator()
                else:
                    # 添加菜单项
                    m.add_command(label=label,image=None,
                        command=tm[label][1], compound=LEFT)
    # 生成所有需要的图标
    def init_icons(self):
        pass
        # self.master.filenew_icon = PhotoImage(name='E:/pydoc/tkinter/images/filenew.png')
        # self.master.fileopen_icon = PhotoImage(name='E:/pydoc/tkinter/images/fileopen.png')
        # self.master.save_icon = PhotoImage(name='E:/pydoc/tkinter/images/save.png')
        # self.master.saveas_icon = PhotoImage(name='E:/pydoc/tkinter/images/saveas.png')
        # self.master.signout_icon = PhotoImage(name='E:/pydoc/tkinter/images/signout.png')
    # 新建项目
    def new_project(self):
        self.new_path()

    # 新建路径
    def new_path(self):
        pass

    def open_filename(self):
        self.excel_path = ''
        # 调用askopenfile方法获取打开的文件名
        self.excel_path = filedialog.askopenfilename(title='打开文件',
            filetypes=[('Excel文件', '*.xlsx'), ('Excel 文件', '*.xls')], # 只处理的文件类型
            initialdir=r'G:\测试结果\battery\rP\GITT') # 初始目录
        self.excel_adr.set(self.excel_path)

    def preview(self):
        if (self.example == 0)and(self.excel_path):
            self.example = Gitt(self.excel_path)
            self.example.read_data()
        elif self.excel_path == '':
            messagebox.showinfo(title='警告',message='请检查文件是否存在！')
        elif (self.example)and(self.excel_path):
            pass
        fig, axs = plt.subplots(2,1)
        axs[0].plot(self.example.pristine_data['测试时间/Sec'], self.example.pristine_data['电压/V'], color='#8080c0')
        axs[0].set_xlim(0, int(self.example.pristine_data['测试时间/Sec'].max())+1)
        axs[0].set_xlabel('time (sec)')
        axs[0].set_ylabel('Potential (V)')
        axs[1].plot(self.example.pristine_data['测试时间/Sec'], self.example.pristine_data['电流/mA'], color='#ff8000')
        axs[1].set_xlim(0, int(self.example.pristine_data['测试时间/Sec'].max())+1)
        axs[1].set_xlabel('time (sec)')
        axs[1].set_ylabel('Currents (mA)')
        axs[1].grid(False)
        fig.tight_layout()
        plt.show()

    def work(self):
        if (self.example == 0)and(self.excel_path):
            self.example = Gitt(self.excel_path)
            self.example.read_data()
        elif self.excel_path == '':
            messagebox.showinfo(title='警告',message='请检查文件是否存在！')
        elif (self.example)and(self.excel_path):
            pass
        self.example.cd_divide(self.example.pristine_data)
        self.DROP = int(self.DROPcheck.get())
        self.result['discharge'] = self.example.diffus_fit(self.example.discharge_data, self.tao, self.DROP, self.customize_Constant)
        self.result['charge'] = self.example.diffus_fit(self.example.charge_data, self.tao, self.DROP, self.customize_Constant)

    def UD_plot(self):
        fig, ax = plt.subplots()
        ax.plot(self.result['discharge']['电压/V'], self.result['discharge']['tEsEt'], 'c*-', linewidth=2, label = 'Discharge')
        ax.plot(self.result['charge']['电压/V'], self.result['charge']['tEsEt'], 'mo-',linewidth=2, label = 'Charge')
        ax.set_xlim(0, int(self.result['discharge']['电压/V'].max())+1)
        ax.set_xlabel('Potential (V)')
        ax.set_ylabel('pi/4D/m^2rou^2')
        ax.set_title('Potential-Diffusion Curves')
        ax.legend()
        plt.show()

    def saveToexcel(self):
        if (len(self.result['discharge'])>0) and (len(self.result['charge'])>0):
            save_path = filedialog.asksaveasfilename(title='保存文件', 
            filetypes=[("office Excel", "*.xls")], # 只处理的文件类型
            initialdir='/Users/hsh/Desktop/')
            # writer = pd.ExcelWriter(save_path) 
            with pd.ExcelWriter(save_path+'.xls') as writer:
                pd.concat([self.result['discharge'], self.result['charge']], axis=1).to_excel(writer, sheet_name='GITT analysis')
            # writer.save()
            # writer.close()
        else:
            yon = messagebox.askquestion(title='提示',message='结果为空，是否先进行数据拟合？')

    def tao_set(self):
        # 调用askinteger函数生成一个让用户输入整数的对话框
        self.tao = simpledialog.askinteger('输入弛豫时间(Sec)', '弛豫时间(Sec):',
            initialvalue=self.tao, minvalue=1, maxvalue=36000)

    def massLoad_set(self):
        self.massload = simpledialog.askinteger('输入活性物质载量(mg)', '活性物质载量(mg):',
            initialvalue=self.massload, minvalue=0.01, maxvalue=10000)

    def actArea_set(self):
        self.actarea = simpledialog.askinteger('输入电化学活性面积(cm^2)', '电化学活性面积(cm^2):',
            initialvalue=self.actarea, minvalue=0.01, maxvalue=10000)

    def density_set(self):
        self.density = simpledialog.askinteger('输入活性物质密度(g/cm^3)', '活性物质密度(g/cm^3):',
            initialvalue=self.density, minvalue=0.0001, maxvalue=10000)

    def original_data_preparation(self):
        messagebox.showinfo(title='原始数据准备',message='从LAND导出一个完整GITT充放循环数据，只导出记录表，选择测试时间、电流、电压、比容量，其中时间单位为秒Sec，保存为Excel。')

    def show_help(self):
        messagebox.showinfo(title='关于',message='离子导率由方程给出：\n' +
            'Dion = (4/π)*n*(m/A^2/ρ^2)*1/τ*(ΔEs/ΔEτ)^2\n' + 'n：反应过程参与电子数\n' + 
            'A：电化学活性面积\n' + 'τ：弛豫时间\n' + 'ΔEs：弛豫终压差\n' + 'ΔEτ：脉冲电势差')

    def select_color(self):
        self.rgb = colorchooser.askcolor(parent=self.master, title='选择线条颜色',
            color = 'black')