# gui.py
# 导入 tkinter 库
import tkinter as tk
# 导入 ttk 库
from tkinter import ttk
# 导入 filedialog 库
from tkinter import filedialog
# 导入主要逻辑代码
import main1
# 导入 sys 库
import sys
# 导入 threading 库
import threading
# 导入 configparser 库
import configparser
from tkinter import messagebox

# 定义一个类，用于封装窗口和控件
class GUI:
    # 定义初始化方法
    def __init__(self):
        # 创建一个窗口
        self.window = tk.Tk()
        self.window.title('SQL to Excel')
        self.window.geometry('400x500')

        # 创建一个 ios 风格的主题
        style = ttk.Style()
        style.theme_use('clam')

        # 创建一个标签，提示输入数据库链接信息
        ttk.Label(self.window, text='请输入数据库的链接信息：').place(x=10, y=10)

        # 定义一个列表，存储输入框和标签的名称和位置信息
        entries = [('host', 10), ('user', 40), ('password', 70), ('db', 100), ('port', 130)]

        # 创建一个字典，存储输入框的变量
        self.entry_vars = {}

        # 用一个循环来创建输入框和标签
        for name, y in entries:
            # 创建一个输入框变量，并存储到字典中
            self.entry_vars[name] = tk.StringVar()
            # 创建一个输入框，并绑定变量和位置
            ttk.Entry(self.window, textvariable=self.entry_vars[name]).place(x=150, y=y)
            # 创建一个标签，并显示名称和位置
            ttk.Label(self.window, text=name).place(x=300, y=y)

        # 创建一个按钮，用于触发测试连接函数
        ttk.Button(self.window, text='测试连接', command=self.test_db_connection).grid(row=2, column=1, pady=60,
                                                                                       padx=20)

        # 创建一个标签，提示选择 .sql 文件的路径
        ttk.Label(self.window, text='请选择 .sql 文件的路径：').place(x=10, y=160)

        # 创建一个输入框变量，用于存储 .sql 文件的路径
        self.sql_path_var = tk.StringVar()

        # 创建一个输入框，并绑定变量和位置
        ttk.Entry(self.window, textvariable=self.sql_path_var).place(x=150, y=160)

        # 创建一个按钮，用于触发选择 .sql 文件路径函数
        self.create_button('选择路径', self.select_sql_path, x=300, y=160)
        # 创建一个标签，提示选择 Excel 文件的路径
        ttk.Label(self.window, text='请选择 Excel 文件的路径：').place(x=10, y=200)

        # 创建一个输入框变量，用于存储 Excel 文件的路径
        self.excel_path_var = tk.StringVar()

        # 创建一个输入框，并绑定变量和位置
        ttk.Entry(self.window, textvariable=self.excel_path_var).place(x=150, y=200)

        # 创建一个按钮，用于触发选择 Excel 文件路径函数
        self.create_button('选择路径', self.select_excel_path, x=300, y=200)

        # 创建一个按钮，用于触发执行 .sql 文件函数，并禁用它
        self.execute_button = self.create_button('执行 .sql 文件', self.execute_sql_file, x=150, y=240)
        self.execute_button['state'] = 'disabled'

        # 创建一个文本框，用于显示日志
        self.text = tk.Text(self.window, width=53, height=15)
        self.text.place(x=10, y=280)

        # 调用 redirect_print 函数，传入文本框对象
        self.redirect_print(self.text)

        # 定义一个函数，用于检测数据库是否连接成功，并弹出提示框

    def test_db_connection(self):
        # 获取配置文件中的数据库链接信息，并赋值给输入框变量
        config = configparser.ConfigParser()
        config.read('config.ini')

        for name in self.entry_vars:
            self.entry_vars[name].set(config['db'][name])

        # 获取输入框的值
        host = self.entry_vars['host'].get()
        user = self.entry_vars['user'].get()
        password = self.entry_vars['password'].get()
        database = self.entry_vars['db'].get()
        port = self.entry_vars['port'].get()

        # 调用主要逻辑代码中的一个函数，来检测数据库是否连接成功
        if main1.test_db_connection(host, user, password, database, port):
            # 弹出提示框，显示数据库连接成功
            tk.messagebox.showinfo('测试连接', '数据库连接成功！')
            # 启用执行按钮
            self.execute_button['state'] = 'normal'
        else:
            # 弹出提示框，显示数据库连接失败
            tk.messagebox.showerror('测试连接', '数据库连接失败！')
            # 禁用执行按钮
            self.execute_button['state'] = 'disabled'

    # 定义一个函数，用于弹出文件选择器，并将选择的路径赋值给输入框
    def select_sql_path(self):
        # 弹出文件选择器，只允许选择文件夹
        sql_path = filedialog.askdirectory(title='请选择 .sql 文件的路径')
        # 将选择的路径赋值给输入框
        self.sql_path_var.set(sql_path)

    # 定义一个函数，用于弹出文件选择器，并将选择的路径赋值给输入框
    def select_excel_path(self):
        # 弹出文件选择器，只允许选择文件夹
        excel_path = filedialog.askdirectory(title='请选择 Excel 文件的路径')
        # 将选择的路径赋值给输入框
        self.excel_path_var.set(excel_path)

    # 定义一个函数，用于创建按钮，并返回按钮对象
    def create_button(self, text, command, x, y):
        # 创建一个按钮，并绑定文本，命令和位置
        button = ttk.Button(self.window, text=text, command=command)
        button.place(x=x, y=y)
        # 返回按钮对象
        return button

    # 定义一个函数，用于执行 .sql 文件，并在一个新的线程中运行
    def execute_sql_file(self):
        # 创建一个线程对象，并传入目标函数和参数
        thread = threading.Thread(target=self.run_sql_file)
        # 启动线程
        thread.start()

    # 定义一个函数，用于获取输入框的值并调用主要逻辑代码
    def run_sql_file(self):
        # 获取输入框的值
        host = self.entry_vars['host'].get()
        user = self.entry_vars['user'].get()
        password = self.entry_vars['password'].get()
        database = self.entry_vars['db'].get()
        sql_path = self.sql_path_var.get()
        excel_path = self.excel_path_var.get()
        port = self.entry_vars['port'].get()

        # 调用主要逻辑代码，并传递参数
        main1.execute_sql_to_excel(host, user, password,
                                  database,
                                  sql_path,
                                  excel_path, port)
    # 定义一个函数，用于重定向 print 函数的输出到文本框中
    def redirect_print(self, text):
        # 获取 print 函数的默认输出对象
        default_stdout = sys.stdout
        # 定义一个类，用于实现 write 方法
        class TextWriter:
            def __init__(self, text):
                self.text = text
            def write(self, s):
                # 在文本框中插入字符串
                self.text.insert(tk.END, s)
                # 滚动到最后一行
                self.text.see(tk.END)
        # 创建一个 TextWriter 对象，传入文本框对象
        writer = TextWriter(text)
        # 将 print 函数的输出对象设置为 writer
        sys.stdout = writer

# 创建一个 GUI 对象
gui = GUI()
# 进入主循环
gui.window.mainloop()