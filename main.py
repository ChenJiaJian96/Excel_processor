import re
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from xlrd import *
from xlwt import *
from time import *
from collections import Counter


# 主界面
class MyGUI:
    def __init__(self):
        self.file_name = None
        self.examiner_list = []  # 考核人员名单
        self.data = None  # master实例
        self.init_window = Tk()  # 父布局

        # 标签
        self.log_label = Label(self.init_window, text="日志")
        self.operate_label = Label(self.init_window, text="执行操作")

        # 文本框
        self.log_data_text = Text(self.init_window)  # 日志文本框
        # 按钮
        self.open_file_button = Button(self.init_window, text="打开文件", command=self.open_file)
        self.export_file_button = Button(self.init_window, text="导出文件", command=self.export_file)
        self.exit_button = Button(self.init_window, text="退出系统", command=self.exit_sys)
        # 滚动条
        self.log_scrollbar_y = Scrollbar(self.init_window)

        # 设置窗口属性
        self.set_init_window()
        self.init_window.mainloop()

    # 定义组件放置位置
    def set_init_window(self):
        self.init_window.title("Excel自动化处理工具")  # 指定标题
        self.init_window.geometry("500x250+100+100")  # 指定初始化大小以及出现位置
        # self.init_window.attributes("-alpha", 0.8)  # 指定透明度

        # 设置组件位置范例
        # self.init_data_text.grid(row=1, column=0, rowspan=5, columnspan=5)
        # 设置滚动条范例
        # self.result_data_scrollbar_y.config(command=self.result_data_text.yview)
        # self.result_data_text.config(yscrollcommand=self.result_data_scrollbar_y.set)
        # self.result_data_scrollbar_y.grid(row=1, column=23, rowspan=15, sticky="NS")

        self.log_label.place(relx=0.05, rely=0.05, relwidth=0.65, relheight=0.15)
        self.operate_label.place(relx=0.75, rely=0.05, relwidth=0.2, relheight=0.15)
        self.log_data_text.place(relx=0.05, rely=0.25, relwidth=0.65, relheight=0.7)
        self.open_file_button.place(relx=0.75, rely=0.25, relwidth=0.2, relheight=0.2)
        self.export_file_button.place(relx=0.75, rely=0.5, relwidth=0.2, relheight=0.2)
        self.exit_button.place(relx=0.75, rely=0.75, relwidth=0.2, relheight=0.2)

        self.log_scrollbar_y.config(command=self.log_data_text.yview)
        self.log_data_text.config(yscrollcommand=self.log_scrollbar_y.set)
        self.log_scrollbar_y.place(relx=0.7, rely=0.25, relheight=0.7)

    # 打开文件
    def open_file(self):
        file_name = filedialog.askopenfilename(title='打开文件', filetypes=[('All Files', '*')])
        self.file_name = file_name
        try:
            temp_data = open_workbook(file_name)
        except FileNotFoundError:
            pass
        except XLRDError:
            self.write_log("请打开正确格式的文件")
        else:
            self.data = ExcelMaster(temp_data)
            self.write_log("打开文件：" + file_name)
            self.check_file_integrity()

    # 检查文件完整性
    def check_file_integrity(self):
        self.write_log("开始检查文件完整性")
        flag = 0
        if self.data.col_index('联系人') == -1:
            self.write_log("打开的文件中找不到列：“联系人”，无法导出员工名单")
            flag = 1
        if self.data.col_index('结束代码') == -1:
            self.write_log("打开的文件中找不到列：“结束代码”, 无法计算员工根本解决率")
            flag = 1

        if flag == 0:
            self.write_log("该文件完整，开始选择考勤名单。")
            self.setup_staff_list()
        else:
            self.write_log("文件不完整，建议检查文件完整性后重启系统。")
            self.set_button_disabled()

    # 使按钮失效，无法使用
    def set_button_disabled(self):
        self.open_file_button.config(state=DISABLED)
        self.export_file_button.config(state=DISABLED)

    # 设置员工列表
    def setup_staff_list(self):
        res = self.open_staff_list()
        print(res)
        if res is None:
            self.write_log("抱歉，你并未选择任何员工")
        else:
            self.write_log("当前选择考核的员工为：" + str(res))
            self.examiner_list = res

    # 开始选择员工列表
    def open_staff_list(self):
        name_list = self.data.get_name_list()
        inputDialog = MyDialog(name_list)
        self.init_window.wait_window(inputDialog.rootWindow)  # 这一句很重要！！！
        return inputDialog.result_list

    # 导出文件
    def export_file(self):
        if self.data is None:
            self.write_log("请先打开您需要处理的文件")
        else:
            self.proceed_data()
            self.write_log("导出成功。")

    # 处理数据
    def proceed_data(self):
        self.get_rate_all_solved()
        self.write_log("处理数据")

    # No.4:获取"事件成功解决率"的数据
    def get_rate_all_solved(self):
        temp_dict = self.data.get_name_dict()
        name_dict = {}
        # 仅保存需要考核的员工
        for name in self.examiner_list:
            try:
                name_dict[name] = temp_dict[name]
            except KeyError:
                pass
        for name in name_dict.keys():
            total_num = name_dict[name]  # 事件总数
            cur_num = self.data.get_num_all_solved(name)  # 根本解决事件数
            rate = float(cur_num / total_num)
            score = self.cal_score_all_solved(rate)
            name_dict[name] = [total_num, cur_num, rate, score]
        # 输出结果
        for name in name_dict.keys():
            print(name + str(name_dict[name]))

        wb = Workbook(encoding='ascii')
        ws = wb.add_sheet("员工根据解决率")
        ws.write(0, 0, "员工姓名")
        ws.write(0, 1, "事件完成数")
        ws.write(0, 2, "事件根本解决数")
        ws.write(0, 3, "事件根本解决率")
        ws.write(0, 4, "该项得分")
        x = 1
        y = 0
        for name in name_dict.keys():
            ws.write(x, y, name)
            y = y + 1
            for item in name_dict[name]:
                ws.write(x, y, item)
                y = y + 1
            x = x + 1
            y = 0
        wb.save('C:\\Users\\Charlie\\Desktop\\test.xls')

    # 添加日志
    def write_log(self, msg):  # 日志动态打印
        current_time = self.get_current_time()
        log_msg = str(current_time) + " " + str(msg) + "\n"  # 换行
        self.log_data_text.insert(END, log_msg)

    @staticmethod
    def exit_sys():
        quit()

    @staticmethod
    def get_current_time():
        current_time = strftime('%Y-%m-%d %H:%M:%S', localtime(time()))
        return current_time

    # 计算根本解决率的得分
    def cal_score_all_solved(self, rate):
        if rate >= 0.995:
            return 100
        elif rate >= 0.98:
            return 90
        elif rate >= 0.8:
            return 80
        elif rate >= 0.7:
            return 70
        else:
            return 10 * int(rate / 0.1)


# 弹窗(采用软耦合的方式接收数据）
class MyDialog:
    def __init__(self, name_list):
        self.name_list = name_list  # 传过来的名单
        self.result_list = []  # 需要发出去的名单
        self.rootWindow = Toplevel()
        self.rootWindow.title('设置考勤名单')
        self.rootWindow.geometry("600x300+250+250")
        self.search_text = Entry(self.rootWindow)
        self.name_list_label = Label(self.rootWindow, text="表格名单(点击多选）")
        self.selected_list_label = Label(self.rootWindow, text="选中名单列表")

        self.search_button = Button(self.rootWindow, text="搜索", command=self.search)
        self.add_button = Button(self.rootWindow, text="添加 >>", command=self.add_name)
        self.del_button = Button(self.rootWindow, text="删除 <<", command=self.del_name)
        self.all_del_button = Button(self.rootWindow, text="全部删除", command=self.del_all)
        self.confirm_button = Button(self.rootWindow, text="确认", command=self.ok)

        self.name_list_box = Listbox(self.rootWindow, selectmode=MULTIPLE)  # 表格员工名单
        self.selected_list_box = Listbox(self.rootWindow, selectmode=BROWSE)  # 选中员工名单
        # 弹窗界面
        self.init_ui()

    def init_ui(self):
        self.search_text.place(relx=0.05, rely=0.05, relwidth=0.6, relheight=0.1)
        self.search_button.place(relx=0.7, rely=0.05, relwidth=0.25, relheight=0.1)
        self.name_list_label.place(relx=0.05, rely=0.15, relwidth=0.3, relheight=0.15)
        self.selected_list_label.place(relx=0.65, rely=0.15, relwidth=0.3, relheight=0.15)
        self.name_list_box.place(relx=0.05, rely=0.3, relwidth=0.3, relheight=0.65)
        self.selected_list_box.place(relx=0.65, rely=0.3, relwidth=0.3, relheight=0.65)
        self.add_button.place(relx=0.4, rely=0.3, relwidth=0.2, relheight=0.12)
        self.del_button.place(relx=0.4, rely=0.47, relwidth=0.2, relheight=0.12)
        self.all_del_button.place(relx=0.4, rely=0.64, relwidth=0.2, relheight=0.12)
        self.confirm_button.place(relx=0.4, rely=0.81, relwidth=0.2, relheight=0.12)
        # TODO: 对名单进行初步处理，去除能删掉的脏数据

        # 对名单进行排序，优化用户体验
        # TODO：排序规则存在问题
        self.refresh_name_list()

    def refresh_name_list(self):
        self.name_list_box.delete(0, END)
        try:
            self.name_list.sort()
        except TypeError:
            messagebox.showwarning("表格内容错误", "表格员工列中出现非法内容，导致列表无法自动排序\n"
                                             "——非法字符包括数字、空格等，请自行删除。")

        for item in self.name_list:
            self.name_list_box.insert(END, item)

    def search(self):
        search_text = self.search_text.get()
        print(search_text)
        if search_text:
            search_result = self.fuzzyfinder(search_text, self.name_list)
            print(search_result)
            self.name_list_box.delete(0, END)
            for name in search_result:
                self.name_list_box.insert(END, name)
        else:
            self.refresh_name_list()

    def add_name(self):
        selected_list = self.name_list_box.curselection()
        print("You have added: " + str(selected_list))
        temp_added_list = []
        for pos in selected_list:
            name = self.name_list_box.get(pos)
            if name in self.result_list:
                temp_added_list.append(name)
            else:
                self.result_list.append(name)
                self.selected_list_box.insert(END, name)
        if temp_added_list:
            messagebox.showwarning("添加员工错误", "以下员工已经添加\n" + str(temp_added_list))
        self.name_list_box.selection_clear(0, END)
        print("Result list: " + str(self.result_list))

    def del_name(self):
        selected_pos = self.selected_list_box.curselection()
        print("You have deleted: " + str(selected_pos))
        self.selected_list_box.delete(selected_pos)
        del self.result_list[int(selected_pos[0])]
        print("Result list: " + str(self.result_list))

    def del_all(self):
        self.selected_list_box.delete(0, END)
        self.result_list.clear()

    def ok(self):
        self.rootWindow.destroy()

    def cancel(self):
        self.result_list = None  # 清空弹窗数据
        self.rootWindow.destroy()

    # 模糊搜索
    def fuzzyfinder(self, user_input, collection):
        suggestions = []
        pattern = '.*'.join(user_input)  # Converts 'djm' to 'd.*j.*m'
        regex = re.compile(pattern)  # Compiles a regex.
        for item in collection:
            match = regex.search(item)  # Checks if the current item matches the regex.
            if match:
                suggestions.append(item)
        return suggestions


# 数据类
class ExcelMaster:
    def __init__(self, data):
        self.data = data  # 源文件
        self.table = None  # 保存当前正在处理的表格
        # 初始化表格
        self.excel_table_by_index()

    # index:第index个sheet,入参需要检查
    def excel_table_by_index(self, index=0):
        if self.data is None:
            return "文件为空，无法打开工作表！"
        else:
            self.table = self.data.sheet_by_index(index)

    # 返回表格的员工列表
    def get_name_list(self):
        i = self.col_index('联系人')
        name_dict = Counter(self.table.col_values(i, start_rowx=1, end_rowx=None))
        return list(name_dict.keys())

    # 返回表格的员工完成事件数
    def get_name_dict(self):
        i = self.col_index('联系人')
        name_dict = Counter(self.table.col_values(i, start_rowx=1, end_rowx=None))
        return name_dict

    # 返回员工“根本解决”的事件总数
    def get_num_all_solved(self, name):
        return 1

    # 计算某位员工的“事件平均相应时长”
    def ave_response_time(self):
        result = dict()
        name_list = self.get_name_list()
        name_dict = Counter(name_list)
        for name in name_dict.keys():
            count = name_dict.get(name)

    # 计算两个字符串时间('%Y/%m/%d %H:%M')的时间差: str2 - str1
    # 返回时间间隔, 单位: s
    def minus_time_in_str(self, str1, str2):
        time1 = strptime(str1, '%Y/%m/%d %H:%M')
        time2 = strptime(str2, '%Y/%m/%d %H:%M')
        return mktime(time2) - mktime(time1)

    # 返回列名返回列索引
    def col_index(self, col_name):
        first_col_list = self.table.row_values(0)  # 第一行元素生成列表
        try:
            i = first_col_list.index(col_name)
        except ValueError:
            return -1
        else:
            return i


MyGUI()  # 启动窗口
