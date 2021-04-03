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
from hslrgb import *


class App():
    def __init__(self, master):
        self.master = master
        self.initWidgets()

        self.excel_path = tuple()
        self.examples = []    # 用于保存实例化对象
        self.tao = [20,20,20,20,20,20,20]  # 脉冲时间，单位：分钟min
        self.massload = [1,1,1,1,1,1,1]    # 活性物质载量，单位：毫克mg
        self.actarea = [1,1,1,1,1,1,1]     # 电化学活性面积，单位：平方厘米cm^2
        self.density = [1,1,1,1,1,1,1]     # 活性物质的密度，单位：立方厘米每克cm^3/g
        self.DROP = 0       #  用于选择丢弃ΔEτ的值的位置，等于0时丢弃最后一个值，等于1时丢弃第一个值
        self.customize_Constant = 2   # 用于截取IR降的位置，数值上等于脉冲开始后的点个数
        self.result = {'discharge': pd.DataFrame(), 'charge': pd.DataFrame()}
        self.results = []   # 用于存储拟合结果

        self.hsl = ('blue', '#B03060')
        
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
        ud_button = Button(fm2, text = 'Show All', 
            bd=3, width = 10, height = 1, 
            command = self.all_plot, 
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
                ('另存为', OrderedDict([('CSV', (None, self.saveTocsv)),
                        ('Excel',(None, self.saveToexcel))])),
                ('-1', (None, None)),
                ('退出', (None, self.master.quit)),
                ]),
            OrderedDict([('预览',(None, self.preview)), 
                ('-1',(None, None)),
                ('U-DLi+',(None, self.UlgD_plot)),
                ('Q-DLi+ ',(None, self.QlgD_plot)),
                ('Q-R ',(None, self.QR_plot)),
                ('-2',(None, None)),
                ('Show All', (None, self.all_plot)),
                ('-3',(None, None)),
                ('选择颜色', (None, self.select_color)),
                ]),
            OrderedDict([('输入测试参数',(None,self.input_para)),
                ('-1',(None, None)),
                ('选取IR降位置',(None,self.IR_set)),
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
                            sm.add_command(label=sub_label,image=sub_dict[sub_label][0],
                                command=sub_dict[sub_label][1], compound=LEFT)
                elif label.startswith('-'):
                    # 添加分隔条
                    m.add_separator()
                else:
                    # 添加菜单项
                    m.add_command(label=label,image=tm[label][0],
                        command=tm[label][1], compound=LEFT)
    # 生成所有需要的图标
    def init_icons(self):
        pass
        # self.master.filenew_icon = PhotoImage(file=r"E:\pydoc\gittview2.0\image\filenew.png")
        # self.master.fileopen_icon = PhotoImage(file=r"E:\pydoc\gittview2.0\image\fileopen.png")
        # self.master.save_icon = PhotoImage(file=r"E:\pydoc\gittview2.0\image\save.png")
        # self.master.saveas_icon = PhotoImage(file=r"E:\pydoc\gittview2.0\image\saveas.png")
        # self.master.csv_icon = PhotoImage(file=r"E:\pydoc\gittview2.0\image\csv.png")
        # self.master.xls_icon = PhotoImage(file=r"E:\pydoc\gittview2.0\image\xls.png")
        # self.master.signout_icon = PhotoImage(file=r"E:\pydoc\gittview2.0\image\signout.png")
        # self.master.preview_icon = PhotoImage(file=r"E:\pydoc\gittview2.0\image\view.png")
    # 新建项目
    def new_project(self):
        self.new_path()
        self.tao = [20,20,20,20,20,20,20] 
        self.massload = [1,1,1,1,1,1,1] 
        self.actarea = [1,1,1,1,1,1,1] 
        self.density = [1,1,1,1,1,1,1] 
        self.customize_Constant = 2
        self.hsl = ('blue', '#B03060')

    # 新建路径
    def new_path(self):
        self.excel_path = tuple()
        self.excel_adr.set('')
        self.DROP = 0
        self.examples = []
        self.result = {'discharge': pd.DataFrame(), 'charge': pd.DataFrame()}
        self.results = []

    def open_filename(self):
        self.excel_path = ()
        # 调用askopenfile方法获取打开的文件名
        self.excel_path = filedialog.askopenfilenames(title='选择一个或多个excel文件',
            filetypes=[('Excel文件', '*.xlsx'), ('Excel 文件', '*.xls')], # 只处理的文件类型
            initialdir=r'C:/Users/Administrator/Desktop') # 初始目录
        self.excel_adr.set(self.excel_path)

    def preview(self):
        if (len(self.examples) == 0)and(self.excel_path):
            for data_path in self.excel_path:
                example = Gitt(data_path)
                example.read_data()
                self.examples.append(example)
        elif len(self.excel_path) == 0:
            messagebox.showinfo(title='警告',message='请检查文件是否存在！')
        elif (self.examples)and(self.excel_path):
            pass
        fig, axs = plt.subplots(2,1)
        color0=self.hsl[1]
        color1=self.loop_pick_color(color0, 1.5)
        for i, example in enumerate(self.examples):
            color00 = self.loop_pick_color(color0, i)
            color11 = self.loop_pick_color(color1, i)
            axs[0].plot(example.pristine_data['测试时间/Sec'], example.pristine_data['电压/V'], 
            color=color00, linewidth = 1, alpha=0.5, label=self.excel_path[i].split('/')[-1][:15])
            axs[1].plot(example.pristine_data['测试时间/Sec'], example.pristine_data['电流/mA'], 
            color=color11, linewidth = 1, alpha=0.5, label=self.excel_path[i].split('/')[-1][:15])
        axs[0].set_xlabel('time (sec)')
        axs[0].set_ylabel('Potential (V)')
        axs[1].set_xlabel('time (sec)')
        axs[1].set_ylabel('Currents (mA)')
        axs[1].grid(False)
        for ax in axs.flat:
            ax.legend()
        fig.tight_layout()
        plt.show()

    def work(self):
        if (len(self.examples) == 0)and(self.excel_path):
            for data_path in self.excel_path:
                example = Gitt(data_path)
                example.read_data()
                self.examples.append(example)
        elif len(self.excel_path) == 0:
            messagebox.showinfo(title='警告',message='请检查文件是否存在！')
        elif (self.examples)and(self.excel_path):
            pass
        self.DROP = int(self.DROPcheck.get())
        for example, t, m, a, p in zip(self.examples, self.tao, self.massload, self.actarea, self.density):
            example.cd_divide(example.pristine_data)

            self.result['discharge'] = example.diffus_fit(example.discharge_data, t, 
                m, a, p, self.DROP, self.customize_Constant)
            self.result['charge'] = example.diffus_fit(example.charge_data, t, 
                m, a, p, self.DROP, self.customize_Constant)

            self.results.append(self.result.copy())

    def UlgD_plot(self):
        try:
            color0=self.hsl[1]
            color1=self.loop_pick_color(color0, 1.5)
            fig, ax = plt.subplots()
            for i, excel_path, result in zip(range(0,len(self.excel_path)), self.excel_path, self.results):
                color00 = self.loop_pick_color(color0, i)
                color11 = self.loop_pick_color(color1, i)
                ax.plot(result['discharge']['电压/V'], np.log10(result['discharge'][r'D/cm\+(2)·s\+(-1)']), 
                color=color00, marker='*', linestyle='solid', linewidth=2, alpha=0.5, label = excel_path.split('/')[-1][:15]+'-Discharge')
                ax.plot(result['charge']['电压/V'], np.log10(result['charge'][r'D/cm\+(2)·s\+(-1)']), 
                color=color11, marker='o', linestyle='solid', linewidth=2, alpha=0.5, label = excel_path.split('/')[-1][:15]+'-Charge')
            ax.set_xlabel('Potential (V)')
            ax.set_ylabel('log10(D) (cm^2/s)')
            ax.legend()
            plt.show()
        except KeyError:
            messagebox.showinfo(title='警告',message='请检查输入文件的有效性！')

    def QlgD_plot(self):
        try:
            color0=self.hsl[1]
            color1=self.loop_pick_color(color0, 1.5)
            fig, ax = plt.subplots()
            for i, excel_path, result in zip(range(0,len(self.excel_path)), self.excel_path, self.results):
                color00 = self.loop_pick_color(color0, i)
                color11 = self.loop_pick_color(color1, i)
                ax.plot(result['discharge']['比容量/mAh/g'], np.log10(result['discharge'][r'D/cm\+(2)·s\+(-1)']), 
                color=color00, marker='*', linestyle='solid', linewidth=2, alpha=0.5, label = excel_path.split('/')[-1][:15]+'-Discharge')
                ax.plot(result['charge']['比容量/mAh/g'], np.log10(result['charge'][r'D/cm\+(2)·s\+(-1)']), 
                color=color11, marker='o', linestyle='solid', linewidth=2, alpha=0.5, label = excel_path.split('/')[-1][:15]+'-Charge')
            # ax.set_xlim(0, int(self.result['discharge']['比容量/mAh/g'].max())*1.1)
            ax.set_xlabel('比容量 (mAh/g)')
            ax.set_ylabel('log10(D) (cm^2/s)')
            # ax.set_title(self.excel_path.split('/')[-1])
            ax.legend()
            plt.show()
        except KeyError:
            messagebox.showinfo(title='警告',message='请检查输入文件的有效性！')

    def QR_plot(self):
        try:
            color0=self.hsl[1]
            color1=self.loop_pick_color(color0, 1.5)
            fig, ax = plt.subplots(1,2)
            for i, excel_path, result in zip(range(0,len(self.excel_path)), self.excel_path, self.results):
                color00 = self.loop_pick_color(color0, i)
                color11 = self.loop_pick_color(color1, i)
                ax[0].plot(result['discharge']['Capacity(R)/mAh/g'], result['discharge']['Reaction resistance/Ohm/g'], 
                color=color00, marker='*', linestyle='solid', linewidth=2, alpha=0.5, label = excel_path.split('/')[-1][:15]+'-Discharge')
                ax[1].plot(result['charge']['Capacity(R)/mAh/g'], result['charge']['Reaction resistance/Ohm/g'], 
                color=color11, marker='o', linestyle='solid', linewidth=2, alpha=0.5, label = excel_path.split('/')[-1][:15]+'-Charge')
            ax[0].set_xlabel('Capacity(R) (mAh/g)')
            ax[0].set_ylabel('Reaction resistance (Ohm/g)')
            ax[0].set_title('Discharge')
            ax[1].set_xlabel('Capacity(R) (mAh/g)')
            ax[1].set_ylabel('Reaction resistance (Ohm/g)')
            ax[1].set_title('Charge')
            ax[0].legend()
            ax[1].legend()
            plt.show()
        except KeyError:
            messagebox.showinfo(title='警告',message='请检查输入文件的有效性！')

    def all_plot(self):
        try:
            # color0 = '#8080c0'
            # color1 = '#ff8000'
            # color2 = '#B03060'
            # color3 = '#3D59AB'
            # color4 = '#DA70D6'
            # color5 = '#7B68EE'
            color0 = self.hsl[1]
            color1 = self.loop_pick_color(color0, 2.39)
            color2 = self.loop_pick_color(color0, 3.8)
            color3 = self.loop_pick_color(color0, 4.2)
            color4 = self.loop_pick_color(color0, 5.6)
            color5 = self.loop_pick_color(color0, 7.3)
            fig, axs = plt.subplots(2, 2)
            for i, excel_path, example, result in zip(range(0,len(self.excel_path)), self.excel_path, self.examples, self.results):
                color00 = self.loop_pick_color(color0, i)
                color11 = self.loop_pick_color(color1, i)
                color22 = self.loop_pick_color(color2, i)
                color33 = self.loop_pick_color(color3, i)
                color44 = self.loop_pick_color(color4, i)
                color55 = self.loop_pick_color(color5, i)
                axs[0, 0].plot(example.discharge_data['测试时间/Sec'], example.discharge_data['电压/V'], label=excel_path.split('/')[-1][:15], 
                color=color00, alpha=0.5)
                axs[0, 0].plot(example.charge_data['测试时间/Sec'], example.charge_data['电压/V'], label=None, 
                color=color00, alpha=0.5)
                axs[0, 1].plot(result['discharge']['电压/V'], np.log10(result['discharge'][r'D/cm\+(2)·s\+(-1)']), 
                color=color11, marker='*', linestyle='solid', linewidth=2, alpha=0.5, label = excel_path.split('/')[-1][:15]+'-Discharge')
                axs[0, 1].plot(result['charge']['电压/V'], np.log10(result['charge'][r'D/cm\+(2)·s\+(-1)']), 
                color=color33, marker='o', linestyle='solid', linewidth=2, alpha=0.5, label = excel_path.split('/')[-1][:15]+'-Charge')
                axs[1, 0].plot(result['discharge']['Capacity(R)/mAh/g'], result['discharge']['Reaction resistance/Ohm/g'], 
                color=color44, marker='*', linestyle='solid', linewidth=2, alpha=0.5, label = excel_path.split('/')[-1][:15])
                axs[1, 1].plot(result['charge']['Capacity(R)/mAh/g'], result['charge']['Reaction resistance/Ohm/g'], 
                color=color55, marker='o', linestyle='solid', linewidth=2, alpha=0.5, label = excel_path.split('/')[-1][:15])
            axs[0, 0].set_xlabel('time (sec)')
            axs[0, 0].set_ylabel('Potential (V)')
            axs[0, 1].set_xlabel('Potential (V)')
            axs[0, 1].set_ylabel('log10(D) (cm^2/s)')
            axs[0, 1].set_title('U-lg(D)')
            axs[1, 0].set_xlabel('Capacity(R) (mAh/g)')
            axs[1, 0].set_ylabel('in-situ Reaction resistance (Ohm/g)')
            axs[1, 0].set_title('Discharge Process')
            axs[1, 1].set_title('Charge Process')
            axs[1, 1].set_xlabel('Capacity(R) (mAh/g)')
            axs[1, 1].set_ylabel('')
            for ax in axs.flat:
                ax.legend()
            fig.tight_layout()
            plt.show()
        except KeyError:
            messagebox.showinfo(title='警告',message='请检查输入文件的有效性！')

    def saveToexcel(self):
        if (len(self.result['discharge'])>0) and (len(self.result['charge'])>0):
            save_path = filedialog.asksaveasfilename(title='保存文件', 
            filetypes=[("office Excel", "*.xls")], # 只处理的文件类型
            initialdir='/Users/hsh/Desktop/')
            for path, result in zip(self.excel_path, self.results):
                tit = path.split('/')[-1].split('.')[0]
                with pd.ExcelWriter(save_path+tit+'.xls') as writer:
                    pd.concat([result['discharge'], result['charge']], axis=1
                    ).to_excel(writer, sheet_name='GITT analysis')
        else:
            yon = messagebox.askquestion(title='提示',message='结果为空!')

    def saveTocsv(self):
        if (len(self.result['discharge'])>0) and (len(self.result['charge'])>0):
            save_path = filedialog.asksaveasfilename(title='保存文件', 
                filetypes=[("逗号分隔符文件", "*.csv")], # 只处理的文件类型
                initialdir='/Users/hsh/Desktop/')
            for path, result in zip(self.excel_path, self.results):
                tit = path.split('/')[-1].split('.')[0]
                pd.concat([result['discharge'], result['charge']], axis=1).to_csv(save_path+tit+'.csv')
        else:
            messagebox.showinfo(title='警告',message='结果为空！')

    def input_para(self):
        if self.excel_path:
            kk = TestPara(self.master, self.excel_path, self.tao, self.massload, self.actarea, self.density)
        else:
            messagebox.showinfo(title='警告',message='请检查输入文件的有效性！')
        self.tao = kk.tao
        self.massload = kk.massload
        self.actarea = kk.actArea
        self.density = kk.density

    # 在色环上隔80˚取色,仅改变色相，不改变明度和饱和度
    def loop_pick_color(self, color, i):
        hsl = toHSL(color)
        loop_h = (hsl[0] + 81*i)%360
        picked_color = toRGB([loop_h, hsl[1], hsl[2]])
        return picked_color

    def select_color(self):
        self.hsl = colorchooser.askcolor(parent=self.master, title='选择线条颜色',
            color = '#B03060')

    def IR_set(self):
        self.customize_Constant = simpledialog.askinteger('输入参数', '选取IR降位置',
            initialvalue=self.customize_Constant, minvalue=0, maxvalue=10)

    def original_data_preparation(self):
        messagebox.showinfo(title='原始数据准备',message='从LAND导出一个完整GITT充放循环数据，只导出记录表，选择测试时间、电流、电压、比容量，其中时间单位为秒Sec，保存为Excel。')

    def show_help(self):
        messagebox.showinfo(title='关于',message='离子导率由Fick第二定律导出：\n' +
            'Dion = (4/π)*n*(m/A^2/ρ^2)*1/τ*(ΔEs/ΔEτ)^2\n' + 'n：反应过程参与电子数\n' + 
            'A：电化学活性面积\n' + 'τ：脉冲时间\n' + 'ΔEs：弛豫终压差\n' + 'ΔEτ：脉冲电势差')

    # 自定义对话框类，继承Toplevel------------------------------------------------------------------------------------------
    # 创建弹窗，用于输入各测试数据文件下的其他测试参数：脉冲时间、质量载量、电化学活性面积和活性物质密度
class TestPara(Toplevel):
    # 定义构造方法
    def __init__(self, parent, excel_path, tao, massload, actarea, density, title = '输入各数据下对应测试参数', modal=False):
        Toplevel.__init__(self, parent)
        self.transient(parent)
        # 设置标题
        if title: self.title(title)
        self.parent = parent
        self.excel_path = excel_path
        self.tao = tao
        self.massload = massload
        self.actarea = actarea
        self.density = density
        # 创建对话框的主体内容
        frame = Frame(self)
        # 调用init_widgets方法来初始化对话框界面
        self.initial_focus = self.init_widgets(frame)
        frame.pack(padx=5, pady=5)
        # 调用init_buttons方法初始化对话框下方的按钮
        self.init_buttons()
        # 根据modal选项设置是否为模式对话框
        if modal: self.grab_set()
        if not self.initial_focus:
            self.initial_focus = self
        # 为"WM_DELETE_WINDOW"协议使用self.cancel_click事件处理方法
        self.protocol("WM_DELETE_WINDOW", self.cancel_click)
        # 根据父窗口来设置对话框的位置
        self.geometry("+%d+%d" % (parent.winfo_rootx()+50,
            parent.winfo_rooty()+50))
        print( self.initial_focus)
        # 让对话框获取焦点
        self.initial_focus.focus_set()
        self.wait_window(self)

    def init_var(self):
        self.tao0, self.mass0, self.area0, self.density0 = DoubleVar(), DoubleVar(), DoubleVar(), DoubleVar()
        self.tao1, self.mass1, self.area1, self.density1 = DoubleVar(), DoubleVar(), DoubleVar(), DoubleVar()
        self.tao2, self.mass2, self.area2, self.density2 = DoubleVar(), DoubleVar(), DoubleVar(), DoubleVar()
        self.tao3, self.mass3, self.area3, self.density3 = DoubleVar(), DoubleVar(), DoubleVar(), DoubleVar()
        self.tao4, self.mass4, self.area4, self.density4 = DoubleVar(), DoubleVar(), DoubleVar(), DoubleVar()
        self.tao5, self.mass5, self.area5, self.density5 = DoubleVar(), DoubleVar(), DoubleVar(), DoubleVar()
        self.tao6, self.mass6, self.area6, self.density6 = DoubleVar(), DoubleVar(), DoubleVar(), DoubleVar()
        self.tao_v = [self.tao0, self.tao1, self.tao2, self.tao3, self.tao4, self.tao5, self.tao6]
        self.massload_v = [self.mass0, self.mass1, self.mass2, self.mass3, self.mass4, self.mass5, self.mass6]
        self.actArea_v = [self.area0, self.area1, self.area2, self.area3, self.area4, self.area5, self.area6]
        self.density_v = [self.density0, self.density1, self.density2, self.density3, self.density4, self.density5, self.density6]
        for t, m, a, p, tt, mm, aa, pp in zip(self.tao_v, self.massload_v, self.actArea_v, self.density_v, 
                              self.tao, self.massload, self.actarea, self.density):
            t.set(tt)
            m.set(mm)
            a.set(aa)
            p.set(pp)

    # 通过该方法来创建自定义对话框的内容
    def init_widgets(self, master):
        # 创建第一个容器
        fm1 = Frame(master)
        # 该容器放在左边排列
        fm1.pack(side=TOP, fill=BOTH, expand=NO)
        Label(fm1, font=('StSong', 10, 'bold'), 
            text='                                 ').pack(side=LEFT, ipadx=5, ipady=5, padx=15, pady=10)
        # for para_tit in ['脉冲时间τ(min)', '活性物质载量(mg)', '电化学活性面积(cm^2)', '活性物质密度(g/cm^3)']:
        for tit in self.excel_path:
            Label(fm1, font=('StSong', 10, 'bold'), text=tit.split('/')[-1][:15]).pack(side=LEFT, padx=10, pady=10)

        self.init_var()
        self.ll = len(self.excel_path)

        fm2 = Frame(master)
        fm2.pack(side=TOP, fill=BOTH, expand=NO)
        Label(fm2, font=('StSong', 10, 'bold'), text='       脉冲时间τ(min)       '
        ).pack(side=LEFT, ipadx=5, ipady=5, padx=5, pady=10)
        for t in self.tao_v[:self.ll]:
            try:
                ttk.Entry(fm2, textvariable=t,
                    width=3,
                    font=('StSong', 10, 'bold'),
                    foreground='#8080c0').pack(side=LEFT, ipadx=5, ipady=5, padx=49, pady=10)
            except IndexError:
                continue

        fm3 = Frame(master)
        fm3.pack(side=TOP, fill=BOTH, expand=NO)
        Label(fm3, font=('StSong', 10, 'bold'), text='    活性物质载量m(mg)   '
        ).pack(side=LEFT, ipadx=5, ipady=5, padx=5, pady=10)
        for m in self.massload_v[:self.ll]:
            try:
                ttk.Entry(fm3, textvariable=m,
                    width=3,
                    font=('StSong', 10, 'bold'),
                    foreground='#8080c0').pack(side=LEFT, ipadx=5, ipady=5, padx=49, pady=10)
            except IndexError:
                continue

        fm4 = Frame(master)
        fm4.pack(side=TOP, fill=BOTH, expand=NO)
        Label(fm4, font=('StSong', 10, 'bold'), text='电化学活性面积A(cm^2) '
        ).pack(side=LEFT, ipadx=5, ipady=5, padx=5, pady=10)
        for a in self.actArea_v[:self.ll]:
            try:
                ttk.Entry(fm4, textvariable=a,
                    width=3,
                    font=('StSong', 10, 'bold'),
                    foreground='#8080c0').pack(side=LEFT, ipadx=5, ipady=5, padx=49, pady=10)
            except IndexError:
                continue

        fm5 = Frame(master)
        fm5.pack(side=TOP, fill=BOTH, expand=NO)
        Label(fm5, font=('StSong', 10, 'bold'), text='活性物质密度ρ(g/cm^3) '
        ).pack(side=LEFT, ipadx=5, ipady=5, padx=5, pady=10)
        for p in self.density_v[:self.ll]:
            try:
                ttk.Entry(fm5, textvariable=p,
                    width=3,
                    font=('StSong', 10, 'bold'),
                    foreground='#8080c0').pack(side=LEFT, ipadx=5, ipady=5, padx=49, pady=10)
            except IndexError:
                continue

    # 通过该方法来创建对话框下方的按钮框
    def init_buttons(self):
        f = Frame(self)
        # 创建"确定"按钮,位置绑定self.ok_click处理方法
        w = Button(f, text="确定", width=10, command=self.ok_click, default=ACTIVE)
        w.pack(side=LEFT, padx=5, pady=5)
        # 创建"取消"按钮,位置绑定self.cancel_click处理方法
        w = Button(f, text="取消", width=10, command=self.cancel_click)
        w.pack(side=LEFT, padx=5, pady=5)
        self.bind("<Return>", self.ok_click)
        self.bind("<Escape>", self.cancel_click)
        f.pack()
    # 该方法可对用户输入的数据进行校验
    def validate(self):
        # 可重写该方法
        return True
    # 该方法可处理用户输入的数据
    def process_input(self):
        self.tao =[]
        self.massload = []
        self.actArea = []
        self.density =[]
        for t, m, a, p in zip(self.tao_v[:self.ll], self.massload_v[:self.ll], self.actArea_v[:self.ll], self.density_v[:self.ll]):
            self.tao.append(t.get())
            self.massload.append(m.get())
            self.actArea.append(a.get())
            self.density.append(p.get())

    def ok_click(self, event=None):
        # print('确定')
        # 如果不能通过校验，让用户重新输入
        if not self.validate():
            self.initial_focus.focus_set()
            return
        self.withdraw()
        self.update_idletasks()
        # 获取用户输入数据
        self.process_input()
        # 将焦点返回给父窗口
        self.parent.focus_set()
        # 销毁自己
        self.destroy()
    def cancel_click(self, event=None):
        # print('取消')
        # 将焦点返回给父窗口
        self.parent.focus_set()
        # 销毁自己
        self.destroy()
