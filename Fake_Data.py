import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog, Listbox, MULTIPLE
import numpy as np
import random
import pandas as pd
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib

matplotlib.use("TkAgg")


plt.rcParams["font.family"] = ["SimHei", "Microsoft YaHei", "Arial", "sans-serif"]
plt.rcParams["axes.unicode_minus"] = True

# 姓名数据库
LAST_NAMES = ["赵", "钱", "孙", "李", "周", "吴", "郑", "王", "冯", "陈",
              "褚", "卫", "蒋", "沈", "韩", "杨", "朱", "秦", "尤", "许",
              "何", "吕", "施", "张", "孔", "曹", "严", "华", "金", "魏",
              "陶", "姜", "戚", "谢", "邹", "喻", "柏", "水", "窦", "章",
              "云", "苏", "潘", "葛", "奚", "范", "彭", "郎", "鲁", "韦",
              "昌", "马", "苗", "凤", "花", "方", "俞", "任", "袁", "柳",
              "鲍", "史", "唐", "费", "廉", "岑", "薛", "雷", "贺", "倪"]

FIRST_NAMES_1 = ["伟", "芳", "娜", "敏", "静", "强", "磊", "军", "洋",
                 "勇", "艳", "杰", "涛", "明", "超", "霞", "平", "刚",
                 "桂", "琴", "辉", "琳", "晶", "华", "梅", "红", "英",
                 "文", "斌", "婷", "宇", "浩", "鹏", "佳", "欣", "雨", "轩"]

FIRST_NAMES_2 = ["秀英", "秀兰", "建华", "冬梅", "建国", "志强", "丽娟",
                 "桂英", "秀荣", "海燕", "红梅", "丽娜", "建军", "婷婷",
                 "宇轩", "雨桐", "思远", "佳琪", "梦琪", "浩然", "子轩",
                 "雨欣", "一诺", "欣怡", "子涵", "梓涵", "晨曦", "雨泽"]


class DataGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Fake_Data 制作人:按时吃饭睡觉的小吴")

        # 设置窗口尺寸
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_width = int(screen_width * 0.85)
        window_height = int(screen_height * 0.85)
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(1000, 600)

        # 数据存储变量
        self.data = pd.DataFrame()
        self.current_columns = []
        self.numeric_columns = []
        self.generated_names = set()
        self.table_fig = None
        self.table_ax = None
        self.table_canvas = None

        # 主框架
        self.main_frame = ttk.Frame(root, padding=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # 顶部工具栏区域
        self.toolbar_frame = ttk.LabelFrame(self.main_frame, text="数据操作", padding=5)
        self.toolbar_frame.pack(fill=tk.X, pady=(0, 5))

        # 数据生成工具区
        self.generation_frame = ttk.LabelFrame(self.toolbar_frame, text="数据生成", padding=5)
        self.generation_frame.pack(side=tk.TOP, fill=tk.X, expand=True, padx=(0, 5), pady=(0, 5))

        # 姓名生成控件
        name_gen_frame = ttk.Frame(self.generation_frame)
        name_gen_frame.pack(side=tk.TOP, fill=tk.X, padx=3, pady=3)

        ttk.Label(name_gen_frame, text="生成人数:").pack(side=tk.LEFT, padx=3, pady=3)
        self.name_count = tk.StringVar(value="20")
        ttk.Entry(name_gen_frame, textvariable=self.name_count, width=6).pack(side=tk.LEFT, padx=3, pady=3)

        ttk.Label(name_gen_frame, text="列名:").pack(side=tk.LEFT, padx=3, pady=3)
        self.name_column = tk.StringVar(value="姓名")
        ttk.Entry(name_gen_frame, textvariable=self.name_column, width=10).pack(side=tk.LEFT, padx=3, pady=3)

        ttk.Button(name_gen_frame, text="生成姓名", command=self.generate_names).pack(side=tk.LEFT, padx=8, pady=3)

        # 数字生成控件
        number_gen_frame = ttk.Frame(self.generation_frame)
        number_gen_frame.pack(side=tk.TOP, fill=tk.X, padx=3, pady=3)

        ttk.Label(number_gen_frame, text="生成数量:").pack(side=tk.LEFT, padx=3, pady=3)
        self.number_count = tk.StringVar(value="20")
        ttk.Entry(number_gen_frame, textvariable=self.number_count, width=6).pack(side=tk.LEFT, padx=3, pady=3)

        ttk.Label(number_gen_frame, text="列名:").pack(side=tk.LEFT, padx=3, pady=3)
        self.number_column = tk.StringVar(value="数据")
        ttk.Entry(number_gen_frame, textvariable=self.number_column, width=10).pack(side=tk.LEFT, padx=3, pady=3)

        ttk.Label(number_gen_frame, text="生成方式:").pack(side=tk.LEFT, padx=3, pady=3)
        self.number_method = tk.StringVar(value="range")
        method_frame = ttk.Frame(number_gen_frame)
        method_frame.pack(side=tk.LEFT, padx=3, pady=3)

        ttk.Radiobutton(method_frame, text="区间", variable=self.number_method, value="range").pack(side=tk.LEFT,
                                                                                                    padx=1)
        ttk.Radiobutton(method_frame, text="正态", variable=self.number_method, value="normal").pack(side=tk.LEFT,
                                                                                                     padx=1)
        ttk.Radiobutton(method_frame, text="均匀", variable=self.number_method, value="uniform").pack(side=tk.LEFT,
                                                                                                      padx=1)

        ttk.Button(number_gen_frame, text="生成数字", command=self.generate_numbers).pack(side=tk.LEFT, padx=8, pady=3)

        # 参数设置区域 - 移至生成数量下方
        self.param_frame = ttk.LabelFrame(self.generation_frame, text="参数设置", padding=5)
        self.param_frame.pack(fill=tk.X, pady=(5, 3))

        # 生成参数区域 - 排列更紧凑
        self.gen_param_frame = ttk.Frame(self.param_frame)
        self.gen_param_frame.pack(fill=tk.X)

        self.range_frame = ttk.Frame(self.gen_param_frame)
        ttk.Label(self.range_frame, text="最小值:").pack(side=tk.LEFT, padx=2, pady=2)
        self.min_val = tk.StringVar(value="0")
        ttk.Entry(self.range_frame, textvariable=self.min_val, width=8).pack(side=tk.LEFT, padx=2, pady=2)

        ttk.Label(self.range_frame, text="最大值:").pack(side=tk.LEFT, padx=2, pady=2)
        self.max_val = tk.StringVar(value="100")
        ttk.Entry(self.range_frame, textvariable=self.max_val, width=8).pack(side=tk.LEFT, padx=2, pady=2)

        self.normal_frame = ttk.Frame(self.gen_param_frame)
        ttk.Label(self.normal_frame, text="均值:").pack(side=tk.LEFT, padx=2, pady=2)
        self.mean_val = tk.StringVar(value="50")
        ttk.Entry(self.normal_frame, textvariable=self.mean_val, width=8).pack(side=tk.LEFT, padx=2, pady=2)

        ttk.Label(self.normal_frame, text="标准差:").pack(side=tk.LEFT, padx=2, pady=2)
        self.std_val = tk.StringVar(value="10")
        ttk.Entry(self.normal_frame, textvariable=self.std_val, width=8).pack(side=tk.LEFT, padx=2, pady=2)

        self.uniform_frame = ttk.Frame(self.gen_param_frame)
        ttk.Label(self.uniform_frame, text="均匀分布(0到1之间)").pack(side=tk.LEFT, padx=2, pady=2)

        self.type_frame = ttk.Frame(self.gen_param_frame)
        ttk.Label(self.type_frame, text="数据类型:").pack(side=tk.LEFT, padx=2, pady=2)
        self.data_type = tk.StringVar(value="float")
        ttk.Radiobutton(self.type_frame, text="整数", variable=self.data_type, value="int").pack(side=tk.LEFT, padx=1)
        ttk.Radiobutton(self.type_frame, text="浮点数", variable=self.data_type, value="float").pack(side=tk.LEFT,
                                                                                                     padx=1)

        # 数据处理工具区
        self.processing_frame = ttk.LabelFrame(self.toolbar_frame, text="数据处理", padding=5)
        self.processing_frame.pack(side=tk.TOP, fill=tk.X, expand=True, padx=(5, 0), pady=(5, 0))

        # 误差设置控件
        error_frame = ttk.Frame(self.processing_frame)
        error_frame.pack(side=tk.TOP, fill=tk.X, padx=3, pady=3)

        ttk.Label(error_frame, text="误差强度 (%):").pack(side=tk.LEFT, padx=3, pady=3)
        self.error_strength = tk.StringVar(value="5")
        ttk.Entry(error_frame, textvariable=self.error_strength, width=6).pack(side=tk.LEFT, padx=3, pady=3)

        ttk.Label(error_frame, text="应用到列:").pack(side=tk.LEFT, padx=3, pady=3)
        self.error_column = tk.StringVar()
        self.error_column_combo = ttk.Combobox(error_frame, textvariable=self.error_column, width=12, state="readonly")
        self.error_column_combo.pack(side=tk.LEFT, padx=3, pady=3)

        ttk.Button(error_frame, text="添加随机误差", command=self.add_random_error).pack(side=tk.LEFT, padx=8, pady=3)

        # 数据管理按钮
        manage_frame = ttk.Frame(self.processing_frame)
        manage_frame.pack(side=tk.TOP, fill=tk.X, padx=3, pady=3)
        manage_frame.pack_configure(anchor=tk.CENTER)

        ttk.Button(manage_frame, text="清空数据", command=self.clear_data).pack(side=tk.LEFT, padx=4)
        ttk.Button(manage_frame, text="重命名列", command=self.rename_column).pack(side=tk.LEFT, padx=4)
        ttk.Button(manage_frame, text="删除列", command=self.delete_column).pack(side=tk.LEFT, padx=4)
        ttk.Button(manage_frame, text="导出数据", command=self.export_data).pack(side=tk.LEFT, padx=4)

        # 绘图参数区域
        self.plot_param_frame = ttk.LabelFrame(self.main_frame, text="绘图参数", padding=10)
        self.plot_param_frame.pack(fill=tk.X, pady=(0, 10))

        plot_left_frame = ttk.Frame(self.plot_param_frame)
        plot_left_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        columns_frame = ttk.Frame(plot_left_frame)
        columns_frame.pack(side=tk.LEFT, padx=5, pady=5)

        ttk.Label(columns_frame, text="选择数据列 (可多选):").pack(anchor=tk.W, padx=5, pady=(0, 5))
        self.plot_columns_listbox = Listbox(columns_frame, selectmode=MULTIPLE, width=20, height=5)
        self.plot_columns_listbox.pack(side=tk.LEFT, padx=5, pady=5)

        scrollbar = ttk.Scrollbar(columns_frame, orient=tk.VERTICAL, command=self.plot_columns_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.plot_columns_listbox.config(yscrollcommand=scrollbar.set)

        chart_type_frame = ttk.Frame(plot_left_frame)
        chart_type_frame.pack(side=tk.LEFT, padx=15, pady=5)

        ttk.Label(chart_type_frame, text="图表类型:").pack(anchor=tk.W, padx=5, pady=(0, 5))
        self.plot_type = tk.StringVar(value="line")
        ttk.Radiobutton(chart_type_frame, text="折线图", variable=self.plot_type, value="line").pack(anchor=tk.W,
                                                                                                     padx=5)
        ttk.Radiobutton(chart_type_frame, text="柱状图", variable=self.plot_type, value="bar").pack(anchor=tk.W, padx=5)
        ttk.Radiobutton(chart_type_frame, text="散点图", variable=self.plot_type, value="scatter").pack(anchor=tk.W,
                                                                                                        padx=5)
        ttk.Radiobutton(chart_type_frame, text="直方图", variable=self.plot_type, value="histogram").pack(anchor=tk.W,
                                                                                                          padx=5)
        ttk.Radiobutton(chart_type_frame, text="箱线图", variable=self.plot_type, value="boxplot").pack(anchor=tk.W,
                                                                                                        padx=5)

        plot_right_frame = ttk.Frame(self.plot_param_frame)
        plot_right_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)

        param_grid = ttk.Frame(plot_right_frame)
        param_grid.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(param_grid, text="图表标题:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        self.plot_title = tk.StringVar(value="多列数据图表")
        ttk.Entry(param_grid, textvariable=self.plot_title, width=25).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

        ttk.Label(param_grid, text="X轴标签:").grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
        self.x_label = tk.StringVar(value="X轴")
        ttk.Entry(param_grid, textvariable=self.x_label, width=20).grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

        ttk.Label(param_grid, text="Y轴标签:").grid(row=1, column=2, padx=10, pady=5, sticky=tk.W)
        self.y_label = tk.StringVar(value="Y轴")
        ttk.Entry(param_grid, textvariable=self.y_label, width=20).grid(row=1, column=3, padx=5, pady=5, sticky=tk.W)

        ttk.Label(param_grid, text="X轴间隔:").grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
        self.x_tick = tk.StringVar(value="auto")
        ttk.Entry(param_grid, textvariable=self.x_tick, width=10).grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)

        ttk.Label(param_grid, text="Y轴间隔:").grid(row=2, column=2, padx=10, pady=5, sticky=tk.W)
        self.y_tick = tk.StringVar(value="auto")
        ttk.Entry(param_grid, textvariable=self.y_tick, width=10).grid(row=2, column=3, padx=5, pady=5, sticky=tk.W)

        buttons_frame = ttk.Frame(plot_right_frame)
        buttons_frame.pack(fill=tk.X, pady=5)

        self.grid_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(buttons_frame, text="显示网格", variable=self.grid_var).pack(side=tk.LEFT, padx=10, pady=5)

        ttk.Button(buttons_frame, text="生成图表", command=self.generate_plot).pack(side=tk.LEFT, padx=15, pady=5)
        ttk.Button(buttons_frame, text="保存图表", command=self.save_plot).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(buttons_frame, text="生成表格", command=self.generate_三线表).pack(side=tk.LEFT, padx=5, pady=5)
        # 导出三线表实体按钮
        ttk.Button(buttons_frame, text="导出表格", command=self.export_三线表直接导出).pack(side=tk.LEFT, padx=5,
                                                                                              pady=5)

        # 数据、图表和三线表展示区域 - 并列显示
        self.display_frame = ttk.Frame(self.main_frame)
        self.display_frame.pack(fill=tk.BOTH, expand=True)
        self.display_frame.pack_propagate(False)

        # 数据展示区域（30%宽度）
        self.data_display_frame = ttk.LabelFrame(self.display_frame, text="数据展示", padding=10)
        self.data_display_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5), pady=0)
        self.data_display_frame.pack_propagate(False)
        self.data_display_frame.configure(width=int(window_width * 0.3))

        self.tree = ttk.Treeview(self.data_display_frame)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar_y = ttk.Scrollbar(self.data_display_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

        scrollbar_x = ttk.Scrollbar(self.data_display_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        scrollbar_x.pack(fill=tk.X)

        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        # 图表展示区域（40%宽度）
        self.chart_display_frame = ttk.LabelFrame(self.display_frame, text="图表展示", padding=10)
        self.chart_display_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 5), pady=0)
        self.chart_display_frame.pack_propagate(False)
        self.chart_display_frame.configure(width=int(window_width * 0.4))

        self.fig = plt.Figure(figsize=(8, 6), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_display_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.canvas.draw()

        # 三线表预览区域（30%宽度）
        self.table_display_frame = ttk.LabelFrame(self.display_frame, text="表格预览", padding=10)
        self.table_display_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0), pady=0)
        self.table_display_frame.pack_propagate(False)
        self.table_display_frame.configure(width=int(window_width * 0.3))

        # 初始化三线表区域
        self.init_table_display()

        # 初始化参数
        self.update_param_frame()
        self.number_method.trace_add("write", lambda *args: self.update_param_frame())

    def init_table_display(self):
        """初始化三线表预览区域"""
        self.table_fig = plt.Figure(figsize=(6, 4), dpi=100)
        self.table_ax = self.table_fig.add_subplot(111)
        self.table_ax.axis('off')
        self.table_ax.text(0.5, 0.5, "请选择数据列并点击'生成三线表'",
                           horizontalalignment='center',
                           verticalalignment='center',
                           fontsize=10)

        self.table_canvas = FigureCanvasTkAgg(self.table_fig, master=self.table_display_frame)
        self.table_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.table_canvas.draw()

    def update_param_frame(self):
        for widget in self.gen_param_frame.winfo_children():
            if widget not in [self.range_frame, self.normal_frame, self.uniform_frame, self.type_frame]:
                continue
            widget.pack_forget()
        method = self.number_method.get()
        if method == "range":
            self.range_frame.pack(side=tk.LEFT, padx=5)
        elif method == "normal":
            self.normal_frame.pack(side=tk.LEFT, padx=5)
        elif method == "uniform":
            self.uniform_frame.pack(side=tk.LEFT, padx=5)
        self.type_frame.pack(side=tk.LEFT, padx=10)

    def update_error_columns(self):
        self.error_column_combo['values'] = self.numeric_columns
        if self.numeric_columns:
            self.error_column.set(self.numeric_columns[0])

    def update_plot_columns(self):
        self.plot_columns_listbox.delete(0, tk.END)
        for col in self.numeric_columns:
            self.plot_columns_listbox.insert(tk.END, col)
        if self.numeric_columns:
            self.plot_columns_listbox.selection_set(0)
            self.plot_title.set(f"多列数据图表")

    def generate_names(self):
        try:
            count = int(self.name_count.get())
            if count <= 0 or count > 1000:
                messagebox.showerror("错误", "生成数量必须为1到1000之间的整数")
                return
            total_possible = len(LAST_NAMES) * (len(FIRST_NAMES_1) + len(FIRST_NAMES_2))
            if count > total_possible:
                messagebox.showerror("错误", f"生成数量超过可能的不重复姓名总数({total_possible})")
                return
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数量")
            return
        column_name = self.name_column.get().strip()
        if not column_name:
            column_name = f"姓名{len(self.current_columns) + 1}"
        if column_name in self.current_columns:
            if not messagebox.askyesno("确认", f"列名 '{column_name}' 已存在，是否替换?"):
                i = 1
                new_name = f"{column_name}_{i}"
                while new_name in self.current_columns:
                    i += 1
                    new_name = f"{column_name}_{i}"
                column_name = new_name
            else:
                if not self.data.empty and column_name in self.data.columns:
                    for name in self.data[column_name].dropna():
                        if name in self.generated_names:
                            self.generated_names.remove(name)
        names = []
        max_attempts = 100
        for _ in range(count):
            attempts = 0
            while attempts < max_attempts:
                last_name = random.choice(LAST_NAMES)
                if random.random() < 0.5:
                    first_name = random.choice(FIRST_NAMES_1)
                else:
                    first_name = random.choice(FIRST_NAMES_2)
                full_name = last_name + first_name
                if full_name not in self.generated_names:
                    self.generated_names.add(full_name)
                    names.append(full_name)
                    break
                attempts += 1
            if attempts >= max_attempts:
                messagebox.showerror("错误", "无法生成足够的不重复姓名，请减少生成数量或稍后再试")
                for name in names:
                    self.generated_names.remove(name)
                return
        self.add_column_to_data(column_name, names, is_numeric=False)

    def generate_numbers(self):
        try:
            count = int(self.number_count.get())
            if count <= 0 or count > 1000:
                messagebox.showerror("错误", "生成数量必须为1到1000之间的整数")
                return
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数量")
            return
        column_name = self.number_column.get().strip()
        if not column_name:
            column_name = f"数据{len(self.current_columns) + 1}"
        if column_name in self.current_columns:
            if not messagebox.askyesno("确认", f"列名 '{column_name}' 已存在，是否替换?"):
                i = 1
                new_name = f"{column_name}_{i}"
                while new_name in self.current_columns:
                    i += 1
                    new_name = f"{column_name}_{i}"
                column_name = new_name
        numbers = []
        method = self.number_method.get()
        try:
            if method == "range":
                min_val = float(self.min_val.get())
                max_val = float(self.max_val.get())
                if min_val >= max_val:
                    messagebox.showerror("错误", "最小值必须小于最大值")
                    return
                numbers = np.random.uniform(min_val, max_val, count)
            elif method == "normal":
                mean = float(self.mean_val.get())
                std = float(self.std_val.get())
                if std <= 0:
                    messagebox.showerror("错误", "标准差必须为正数")
                    return
                numbers = np.random.normal(mean, std, count)
            elif method == "uniform":
                numbers = np.random.uniform(0, 1, count)
            if self.data_type.get() == "int":
                numbers = np.round(numbers).astype(int)
            self.add_column_to_data(column_name, numbers.tolist(), is_numeric=True)
        except ValueError as e:
            messagebox.showerror("错误", f"参数输入错误: {str(e)}")
        except Exception as e:
            messagebox.showerror("错误", f"生成数据时出错: {str(e)}")

    def add_column_to_data(self, column_name, data, is_numeric):
        if not self.data.empty:
            current_length = len(self.data)
            new_length = len(data)
            if new_length > current_length:
                data = data[:current_length]
            elif new_length < current_length:
                data += [None] * (current_length - new_length)
        self.data[column_name] = data
        self.current_columns = list(self.data.columns)
        if is_numeric and column_name not in self.numeric_columns:
            self.numeric_columns.append(column_name)
            self.update_error_columns()
            self.update_plot_columns()
        self.update_table_display()

    def update_table_display(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")
            self.tree.column(col, width=0)
        self.tree["columns"] = self.current_columns
        for col in self.current_columns:
            self.tree.heading(col, text=col)
            col_width = min(max(len(col) * 15 + 30, 80), 150)
            self.tree.column(col, width=col_width, anchor=tk.W)
        for i, row in self.data.iterrows():
            values = [str(row[col]) if row[col] is not None else "" for col in self.current_columns]
            self.tree.insert("", tk.END, text=str(i + 1), values=values)

    def add_random_error(self):
        if not self.numeric_columns:
            messagebox.showinfo("提示", "没有可添加误差的数值列")
            return
        column_name = self.error_column.get()
        if not column_name or column_name not in self.numeric_columns:
            messagebox.showerror("错误", "请选择有效的数值列")
            return
        try:
            strength = float(self.error_strength.get())
            if strength <= 0 or strength > 100:
                messagebox.showerror("错误", "误差强度必须为0到100之间的数值")
                return
            strength = strength / 100
            column_data = self.data[column_name].copy()
            valid_data = [x for x in column_data if x is not None]
            if not valid_data:
                messagebox.showinfo("提示", "选中的列没有有效数据")
                return
            data_min = min(valid_data)
            data_max = max(valid_data)
            data_range = data_max - data_min
            if data_range == 0:
                reference_value = abs(valid_data[0]) if valid_data[0] != 0 else 1.0
            else:
                reference_value = data_range
            error_std = strength * reference_value
            new_data = []
            for value in column_data:
                if value is None:
                    new_data.append(None)
                    continue
                noise = np.random.normal(0, error_std)
                new_value = value + noise
                if isinstance(value, int) or (isinstance(value, float) and value.is_integer()):
                    new_value = round(new_value)
                new_data.append(new_value)
            self.data[column_name] = new_data
            self.update_table_display()
            messagebox.showinfo("成功", f"已为 '{column_name}' 列添加随机误差")
        except ValueError as e:
            messagebox.showerror("错误", f"参数输入错误: {str(e)}")
        except Exception as e:
            messagebox.showerror("错误", f"添加误差时出错: {str(e)}")

    def generate_plot(self):
        selected_indices = self.plot_columns_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("提示", "请至少选择一列数据")
            return
        selected_columns = [self.plot_columns_listbox.get(i) for i in selected_indices]
        for col in selected_columns:
            if col not in self.numeric_columns:
                messagebox.showerror("错误", f"列 '{col}' 不是有效的数值列")
                return
        valid_columns = []
        data_dict = {}
        for col in selected_columns:
            data = self.data[col].dropna()
            if len(data) > 0:
                data_dict[col] = data
                valid_columns.append(col)
        if not valid_columns:
            messagebox.showinfo("提示", "选中的列没有有效数据")
            return
        self.ax.clear()
        plot_title = self.plot_title.get()
        x_label = self.x_label.get()
        y_label = self.y_label.get()
        show_grid = self.grid_var.get()
        x_tick = self.x_tick.get().strip().lower()
        y_tick = self.y_tick.get().strip().lower()
        colors = ['b', 'g', 'r', 'c', 'm', 'y', 'k']
        markers = ['o', 's', '^', 'D', 'v', '<', '>']
        line_styles = ['-', '--', '-.', ':']
        plot_type = self.plot_type.get()
        if plot_type == "line":
            for i, (col, data) in enumerate(data_dict.items()):
                color_idx = i % len(colors)
                marker_idx = i % len(markers)
                line_idx = i % len(line_styles)
                self.ax.plot(range(len(data)), data, label=col,
                             color=colors[color_idx],
                             marker=markers[marker_idx],
                             linestyle=line_styles[line_idx],
                             alpha=0.7)
            if not plot_title:
                plot_title = "多列数据折线图"
            if not x_label:
                x_label = "索引"
            if not y_label:
                y_label = "数值"
        elif plot_type == "bar":
            n_cols = len(data_dict)
            bar_width = 0.8 / n_cols
            indices = np.arange(len(next(iter(data_dict.values()))))
            for i, (col, data) in enumerate(data_dict.items()):
                color_idx = i % len(colors)
                positions = indices + i * bar_width
                self.ax.bar(positions, data, width=bar_width, label=col,
                            color=colors[color_idx], alpha=0.7)
            self.ax.set_xticks(indices + bar_width * (n_cols - 1) / 2)
            self.ax.set_xticklabels([str(i) for i in indices])
            if not plot_title:
                plot_title = "多列数据柱状图"
            if not x_label:
                x_label = "索引"
            if not y_label:
                y_label = "数值"
        elif plot_type == "scatter":
            for i, (col, data) in enumerate(data_dict.items()):
                color_idx = i % len(colors)
                marker_idx = i % len(markers)
                self.ax.scatter(range(len(data)), data, label=col,
                                color=colors[color_idx],
                                marker=markers[marker_idx],
                                alpha=0.7)
            if not plot_title:
                plot_title = "多列数据散点图"
            if not x_label:
                x_label = "索引"
            if not y_label:
                y_label = "数值"
        elif plot_type == "histogram":
            for i, (col, data) in enumerate(data_dict.items()):
                color_idx = i % len(colors)
                self.ax.hist(data, bins=min(30, len(data) // 2),
                             alpha=0.5, label=col,
                             color=colors[color_idx], edgecolor='black')
            if not plot_title:
                plot_title = "多列数据直方图"
            if not x_label:
                x_label = "数值"
            if not y_label:
                y_label = "频数"
        elif plot_type == "boxplot":
            data_list = [data for _, data in data_dict.items()]
            self.ax.boxplot(data_list, patch_artist=True,
                            tick_labels=valid_columns,
                            boxprops=dict(facecolor='lightblue', alpha=0.7))
            if not plot_title:
                plot_title = "多列数据箱线图"
            if not y_label:
                y_label = "数值"
            x_label = ""
        self.ax.set_title(plot_title)
        if x_label:
            self.ax.set_xlabel(x_label)
        if y_label:
            self.ax.set_ylabel(y_label)
        if plot_type != "boxplot":
            self.ax.legend()
        try:
            if x_tick != "auto" and x_tick:
                x_interval = float(x_tick)
                if x_interval > 0:
                    x_min, x_max = self.ax.get_xlim()
                    self.ax.set_xticks(np.arange(x_min, x_max + x_interval, x_interval))
            if y_tick != "auto" and y_tick:
                y_interval = float(y_tick)
                if y_interval > 0:
                    y_min, y_max = self.ax.get_ylim()
                    self.ax.set_yticks(np.arange(y_min, y_max + y_interval, y_interval))
        except ValueError:
            messagebox.showwarning("警告", "坐标轴间隔设置无效，将使用自动间隔")
        if show_grid:
            self.ax.grid(True, linestyle='--', alpha=0.7)
        else:
            self.ax.grid(False)
        self.fig.subplots_adjust(
            left=0.1,  # 左边距
            bottom=0.15,  # 底边距（增大以避免标签被截断）
            right=0.9,  # 右边距
            top=0.9  # 顶边距
        )
        self.canvas.draw()

    def save_plot(self):
        if len(self.ax.get_children()) <= 1:
            messagebox.showinfo("提示", "请先生成图表再保存")
            return
        size_str = simpledialog.askstring("图片尺寸",
                                          "请输入图片尺寸 (宽度,高度，单位：英寸)\n例如: 8,6",
                                          initialvalue="8,6")
        if not size_str:
            return
        try:
            width, height = map(float, size_str.split(','))
            if width <= 0 or height <= 0:
                messagebox.showerror("错误", "宽度和高度必须为正数")
                return
        except ValueError:
            messagebox.showerror("错误", "请输入有效的尺寸格式，例如: 8,6")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[
                ("PNG图片", "*.png"),
                ("JPG图片", "*.jpg"),
                ("SVG图片", "*.svg"),
                ("PDF文件", "*.pdf"),
                ("所有文件", "*.*")
            ],
            title="保存图表"
        )
        if not file_path:
            return
        try:
            self.fig.set_size_inches(width, height)
            self.fig.savefig(file_path, dpi=300, bbox_inches='tight')
            self.fig.set_size_inches(8, 6)
            self.canvas.draw()
            messagebox.showinfo("成功", f"图表已成功保存到:\n{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存图表时出错:\n{str(e)}")

    def generate_三线表(self):
        """生成符合学术规范的三线表，严格控制线条、布局、字体"""
        selected_indices = self.plot_columns_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("提示", "请至少选择一列数据生成三线表")
            return
        selected_columns = [self.plot_columns_listbox.get(i) for i in selected_indices]
        for col in selected_columns:
            if col not in self.numeric_columns:
                messagebox.showerror("错误", f"列 '{col}' 不是有效的数值列")
                return

        # 计算均值±标准差
        stats_data = []
        for col in selected_columns:
            data = self.data[col].dropna()
            if len(data) == 0:
                stats_str = "无数据"
            else:
                col_mean = data.mean()
                col_std = data.std()
                stats_str = f"{col_mean:.2f}±{col_std:.2f}"
            stats_data.append(stats_str)

        # 构造表格数据：[列名行, 统计结果行]
        table_content = [selected_columns, stats_data]

        # 调用封装函数绘制学术三线表
        self.draw_academic_table(table_content, "数据统计结果（均值±标准差）")

    def draw_academic_table(self, table_content, title):
        """封装学术三线表绘制逻辑，移除左右框线"""
        # 清空并重置子图
        self.table_ax.clear()
        self.table_ax.axis('off')  # 隐藏默认坐标轴

        # 创建表格：通过 bbox 压缩外间距，让表格更紧凑
        table = self.table_ax.table(
            cellText=table_content,
            colLabels=None,  # 列名直接放第一行，不用单独标签
            cellLoc='center',  # 单元格内容居中
            loc='center',
            bbox=[0.1, 0.1, 0.8, 0.8]  # 缩小表格外间距
        )

        # 表格样式设置：三线表 + 紧凑布局 + 统一字体
        # 1. 线条控制：只保留顶线、栏目线、底线，去除左右边框
        for key, cell in table.get_celld().items():
            row, col = key

            # 重置所有边框为0（清除默认边框）
            #cell.set_linewidth(0)

            # 顶线（第一行上边框）
            if row == 0:
                cell._top_border = 2.0  # 粗线

            # 栏目线（第一行下边框）
            if row == 0:
                cell._bottom_border = 0.7  # 细线
            # 底线（最后一行下边框）
            if row == len(table_content) - 1:
                cell._bottom_border = 2.0  # 粗线

        # 2. 字体与内边距设置
        table.auto_set_font_size(False)
        table.set_fontsize(12)
        table.scale(1.0, 1.5)  # 调整行高

        # 3. 标题设置
        self.table_ax.set_title(title, fontsize=14, y=1.05)

        # 刷新画布
        self.table_fig.tight_layout()
        self.table_canvas.draw()

    def save三线表(self):
        try:
            size_str = simpledialog.askstring("图片尺寸",
                                              "请输入图片尺寸 (宽度,高度，单位：英寸)\n例如: 10,3",
                                              initialvalue="10,3")
            if not size_str:
                return
            width, height = map(float, size_str.split(','))
            if width <= 0 or height <= 0:
                messagebox.showerror("错误", "宽度和高度必须为正数")
                return
        except ValueError:
            messagebox.showerror("错误", "请输入有效的尺寸格式，例如: 10,3")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[
                ("PNG图片", "*.png"),
                ("JPG图片", "*.jpg"),
                ("SVG图片", "*.svg"),
                ("PDF文件", "*.pdf"),  # PDF格式适合学术出版
                ("所有文件", "*.*")
            ],
            title="保存三线表图片"
        )
        if not file_path:
            return

        try:
            original_size = self.table_fig.get_size_inches()
            self.table_fig.set_size_inches(width, height)
            self.table_fig.savefig(file_path, dpi=300, bbox_inches='tight')
            self.table_fig.set_size_inches(original_size)
            self.table_canvas.draw()
            messagebox.showinfo("成功", f"三线表图片已成功保存到:\n{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存三线表图片时出错:\n{str(e)}")

    def export三线表数据(self):
        try:
            # 准备导出数据，只包含列名和均值±标准差
            export_data = self.table_stats_for_export.to_frame().T.reset_index(drop=True)
            row_labels = ["均值±标准差"]
            export_data.insert(0, "统计量", row_labels)  # 更学术化的行标签

            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[
                    ("CSV文件", "*.csv"),
                    ("Excel文件", "*.xlsx"),  # Excel格式适合进一步编辑
                    ("所有文件", "*.*")
                ],
                title="导出三线表数据"
            )
            if not file_path:
                return

            ext = os.path.splitext(file_path)[1].lower()
            if ext == ".csv":
                export_data.to_csv(file_path, index=False, encoding="utf-8-sig")
            elif ext == ".xlsx":
                try:
                    export_data.to_excel(file_path, index=False, engine="openpyxl")
                except ImportError:
                    messagebox.showerror("错误", "导出Excel需安装openpyxl库\n请执行: pip install openpyxl")
                    return
            else:
                export_data.to_csv(file_path, index=False, encoding="utf-8-sig")

            messagebox.showinfo("成功", f"三线表数据已成功导出到:\n{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出三线表数据时出错:\n{str(e)}")

    def export_三线表直接导出(self):
        """直接导出符合学术规范的三线表数据"""
        selected_indices = self.plot_columns_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("提示", "请至少选择一列数据导出三线表")
            return
        selected_columns = [self.plot_columns_listbox.get(i) for i in selected_indices]
        for col in selected_columns:
            if col not in self.numeric_columns:
                messagebox.showerror("错误", f"列 '{col}' 不是有效的数值列")
                return

        # 计算均值±标准差
        stats_row = []
        for col in selected_columns:
            data = self.data[col].dropna()
            if len(data) == 0:
                stats_str = "无数据"
            else:
                col_mean = data.mean()
                col_std = data.std()
                stats_str = f"{col_mean:.2f}±{col_std:.2f}"
            stats_row.append(stats_str)

        # 准备导出数据
        export_data = pd.Series(stats_row, index=selected_columns, name="均值±标准差").to_frame().T.reset_index(
            drop=True)
        row_labels = ["均值±标准差"]
        export_data.insert(0, "统计量", row_labels)  # 学术化的行标签

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[
                ("CSV文件", "*.csv"),
                ("Excel文件", "*.xlsx"),
                ("所有文件", "*.*")
            ],
            title="导出三线表数据"
        )
        if not file_path:
            return

        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext == ".csv":
                export_data.to_csv(file_path, index=False, encoding="utf-8-sig")
            elif ext == ".xlsx":
                try:
                    export_data.to_excel(file_path, index=False, engine="openpyxl")
                except ImportError:
                    messagebox.showerror("错误", "导出Excel需安装openpyxl库\n请执行: pip install openpyxl")
                    return
            else:
                export_data.to_csv(file_path, index=False, encoding="utf-8-sig")

            messagebox.showinfo("成功", f"三线表数据已成功导出到:\n{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出三线表数据时出错:\n{str(e)}")

    def clear_data(self):
        if self.data.empty:
            messagebox.showinfo("提示", "数据已为空")
            return
        if messagebox.askyesno("确认", "确定要清空所有数据吗?"):
            self.generated_names.clear()
            self.data = pd.DataFrame()
            self.current_columns = []
            self.numeric_columns = []
            self.update_error_columns()
            self.update_plot_columns()
            self.update_table_display()
            self.ax.clear()
            self.canvas.draw()
            # 重置三线表预览区域
            self.table_ax.clear()
            self.table_ax.axis('off')
            self.table_ax.text(0.5, 0.5, "请选择数据列并点击'生成三线表'",
                               horizontalalignment='center',
                               verticalalignment='center',
                               fontsize=10)
            self.table_canvas.draw()

    def rename_column(self):
        if not self.current_columns:
            messagebox.showinfo("提示", "没有可重命名的列")
            return
        column = simpledialog.askstring("选择列",
                                        f"请输入要重命名的列名\n当前列名: {', '.join(self.current_columns)}")
        if not column or column not in self.current_columns:
            messagebox.showerror("错误", "无效的列名")
            return
        new_name = simpledialog.askstring("重命名列", f"请输入 '{column}' 的新列名:")
        if not new_name:
            messagebox.showinfo("提示", "未输入新列名")
            return
        if new_name in self.current_columns and new_name != column:
            messagebox.showerror("错误", "新列名已存在")
            return
        is_numeric = column in self.numeric_columns
        self.data = self.data.rename(columns={column: new_name})
        self.current_columns = list(self.data.columns)
        if is_numeric:
            self.numeric_columns.remove(column)
            self.numeric_columns.append(new_name)
            self.update_error_columns()
            self.update_plot_columns()
        self.update_table_display()

    def delete_column(self):
        if not self.current_columns:
            messagebox.showinfo("提示", "没有可删除的列")
            return
        column = simpledialog.askstring("选择列",
                                        f"请输入要删除的列名\n当前列名: {', '.join(self.current_columns)}")
        if not column or column not in self.current_columns:
            messagebox.showerror("错误", "无效的列名")
            return
        if column not in self.numeric_columns:
            if not self.data.empty and column in self.data.columns:
                for name in self.data[column].dropna():
                    if name in self.generated_names:
                        self.generated_names.remove(name)
        if messagebox.askyesno("确认", f"确定要删除 '{column}' 列吗?"):
            if column in self.numeric_columns:
                self.numeric_columns.remove(column)
                self.update_error_columns()
                self.update_plot_columns()
            self.data = self.data.drop(columns=[column])
            self.current_columns = list(self.data.columns)
            self.update_table_display()

    def export_data(self):
        if self.data.empty:
            messagebox.showinfo("提示", "没有可导出的数据")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[
                ("CSV文件", "*.csv"),
                ("文本文件", "*.txt"),
                ("Excel文件", "*.xlsx"),
                ("所有文件", "*.*")
            ],
            title="保存数据"
        )
        if not file_path:
            return
        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext == ".csv":
                self.data.to_csv(file_path, index=False, encoding="utf-8-sig")
            elif ext == ".txt":
                self.data.to_string(file_path, index=False, encoding="utf-8")
            elif ext == ".xlsx":
                try:
                    self.data.to_excel(file_path, index=False)
                except ImportError:
                    messagebox.showerror("错误", "导出为Excel需要安装openpyxl库\n请使用命令: pip install openpyxl")
                    return
            else:
                self.data.to_csv(file_path, index=False, encoding="utf-8-sig")
            messagebox.showinfo("成功", f"数据已成功导出到:\n{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出数据时出错:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DataGeneratorApp(root)
    root.mainloop()
