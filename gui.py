# gui.py
# 导入 tkinter 库
import tkinter as tk
# 导入 ttk 库
from tkinter import ttk
# 导入 filedialog 库
from tkinter import filedialog
# 导入主要逻辑代码
import 工作.main
# 导入 sys 库
import sys

# 创建一个窗口
window = tk.Tk()
window.title('SQL to Excel')
window.geometry('400x500')

# 创建一个 ios 风格的主题
style = ttk.Style()
style.theme_use('clam')

# 创建一个标签
ttk.Label(window, text='请输入数据库的链接信息：').place(x=10, y=10)

# 创建四个输入框
host_var = tk.StringVar()
ttk.Entry(window, textvariable=host_var).place(x=150, y=10)
ttk.Label(window, text='host').place(x=300, y=10)

user_var = tk.StringVar()
ttk.Entry(window, textvariable=user_var).place(x=150, y=40)
ttk.Label(window, text='user').place(x=300, y=40)

password_var = tk.StringVar()
ttk.Entry(window, textvariable=password_var).place(x=150, y=70)
ttk.Label(window, text='password').place(x=300, y=70)

db_var = tk.StringVar()
ttk.Entry(window, textvariable=db_var).place(x=150, y=100)
ttk.Label(window, text='db').place(x=300, y=100)

port_var = tk.StringVar()
ttk.Entry(window, textvariable=port_var).place(x=150, y=130)
ttk.Label(window, text='port').place(x=300, y=130)

# 定义一个函数，用于检测数据库是否连接成功
def test_db_connection():
    # 获取输入框的值
    host = host_var.get()
    user = user_var.get()
    password = password_var.get()
    database = db_var.get()
    port=port_var.get()
    # 调用主要逻辑代码中的一个函数，来检测数据库是否连接成功
    if 工作.main.test_db_connection(host, user, password, database,port):
        # 弹出提示框，显示数据库连接成功
        tk.messagebox.showinfo('测试连接', '数据库连接成功！')
    else:
        # 弹出提示框，显示数据库连接失败
        tk.messagebox.showerror('测试连接', '数据库连接失败！')

# 创建一个按钮，用于触发函数
ttk.Button(window, text='测试连接', command=test_db_connection).grid(row=2, column=1, pady=60,padx=20)
# 创建一个标签
ttk.Label(window, text='请选择 .sql 文件的路径：').place(x=10, y=160)

# 创建一个输入框
sql_path_var = tk.StringVar()
ttk.Entry(window, textvariable=sql_path_var).place(x=150, y=160)

# 定义一个函数，用于弹出文件选择器，并将选择的路径赋值给输入框
def select_sql_path():
    # 弹出文件选择器，只允许选择文件夹
    sql_path = filedialog.askdirectory(title='请选择 .sql 文件的路径')
    # 将选择的路径赋值给输入框
    sql_path_var.set(sql_path)

# 创建一个按钮，用于触发函数
ttk.Button(window, text='浏览', command=select_sql_path).place(x=300, y=160)

# 创建一个标签
ttk.Label(window, text='请选择 excel 输出的路径：').place(x=10, y=200)

# 创建一个输入框
excel_path_var = tk.StringVar()
ttk.Entry(window, textvariable=excel_path_var).place(x=150, y=200)

# 定义一个函数，用于弹出文件选择器，并将选择的路径赋值给输入框
def select_excel_path():
    # 弹出文件选择器，只允许选择文件夹
    excel_path = filedialog.askdirectory(title='请选择 excel 输出的路径')
    # 将选择的路径赋值给输入框
    excel_path_var.set(excel_path)
# 创建一个按钮，用于触发函数
ttk.Button(window, text='浏览', command=select_excel_path).place(x=300, y=200)

# 定义一个函数，用于获取输入框的值并调用主要逻辑代码
def run():
    # 获取输入框的值
    host = host_var.get()
    user = user_var.get()
    password = password_var.get()
    database = db_var.get()
    sql_path = sql_path_var.get()
    excel_path = excel_path_var.get()
    port=port_var.get()
    # 调用主要逻辑代码，并传递参数
    工作.main.execute_sql_to_excel(host, user, password,
                              database,
                              sql_path,
                              excel_path,port)

# 创建一个按钮，用于触发函数
ttk.Button(window,
           text='运行',
           command=run).place(x=150,
                              y=240)

# 创建一个文本框，用于显示日志
text = tk.Text(window, width=53, height=15)
text.place(x=10, y=280)

# 定义一个函数，用于重定向 print 函数的输出到文本框中
def redirect_print(text):
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

# 调用 redirect_print 函数，传入文本框对象
redirect_print(text)

# 进入主循环 fff
window.mainloop()
