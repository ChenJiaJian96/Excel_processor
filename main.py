from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from xlrd import *
from xlwt import *
from time import *
from collections import Counter

# TODO：导出文件加上工号
# 打包exe文件
# pyinstaller -F -w main.py

global option1, option2, option3, option4, option5, option6, option7
option1 = "全部导出"
option2 = "仅导出指定员工的“平均响应时长”情况"
option3 = "仅导出指定员工的“响应超时率”情况"
option4 = "仅导出指定员工的“按时解决率”情况"
option5 = "仅导出指定员工的“根本解决”情况"
option6 = "仅导出指定员工的“平均满意度”情况"
option7 = "仅导出指定员工的“平均解决时长”情况"


# 主界面
class MyGUI:
    def __init__(self):
        self.file_name = None
        self.examiner_list = []  # 考核人员名单
        self.data = None  # master实例
        self.init_window = Tk()  # 父布局
        self.final_wb = Workbook(encoding='ascii')  # 最终导出的文件实例

        # 标签
        self.log_label = Label(self.init_window, text="日志")
        self.operate_label = Label(self.init_window, text="执行操作")

        # 文本框
        self.log_data_text = Text(self.init_window)  # 日志文本框
        # 按钮
        self.open_file_button = Button(self.init_window, text="打开文件", command=self.open_file)
        self.export_file_button = Button(self.init_window, text="导出文件", command=self.export_file)
        self.change_examiners_button = Button(self.init_window, text="修改考核名单", command=self.setup_staff_list)
        self.exit_button = Button(self.init_window, text="退出系统", command=self.exit_sys)
        # 滚动条
        self.log_scrollbar_y = Scrollbar(self.init_window)
        # 图标
        self.more_label = Label(self.init_window, text="...", font="bold, 8")
        self.question_label = Label(self.init_window, text=" ? ", font="bold, 8")
        self.exclamation_label = Button(self.init_window, text=" ! ", font="bold, 8")
        self.bottom_label = Label(self.init_window, text="@CopyRight", font="Arial, 8")
        # 设置窗口属性
        self.set_init_window()
        self.init_window.mainloop()

    # 定义组件放置位置
    def set_init_window(self):
        self.init_window.title("Excel自动化处理工具")  # 指定标题
        self.init_window.geometry("500x265+100+100")  # 指定初始化大小以及出现位置
        # self.init_window.attributes("-alpha", 0.8)  # 指定透明度

        self.log_label.place(relx=0.05, rely=0.05, relwidth=0.6, relheight=0.1)
        self.operate_label.place(relx=0.7, rely=0.05, relwidth=0.2, relheight=0.1)
        self.log_data_text.place(relx=0.05, rely=0.20, relwidth=0.6, relheight=0.75)

        self.open_file_button.place(relx=0.72, rely=0.18, relwidth=0.18, relheight=0.17)
        self.export_file_button.place(relx=0.72, rely=0.38, relwidth=0.18, relheight=0.17)
        self.change_examiners_button.place(relx=0.72, rely=0.58, relwidth=0.18, relheight=0.17)
        self.exit_button.place(relx=0.72, rely=0.78, relwidth=0.18, relheight=0.17)

        self.log_scrollbar_y.config(command=self.log_data_text.yview)
        self.log_data_text.config(yscrollcommand=self.log_scrollbar_y.set)
        self.log_scrollbar_y.place(relx=0.65, rely=0.2, relheight=0.75)

        # 生成右侧提示按钮
        self.more_label.place(relx=0.93, rely=0.7, relwidth=0.03, relheight=0.08)
        self.question_label.place(relx=0.93, rely=0.8, relwidth=0.03, relheight=0.08)
        self.exclamation_label.place(relx=0.93, rely=0.9, relwidth=0.03, relheight=0.08)

        self.bottom_label.place(relx=0.4, rely=0.95, relwidth=0.2, relheight=0.05)

    # 打开文件
    def open_file(self):
        file_name = filedialog.askopenfilename(title='打开文件',
                                               filetypes=[('表格文件', '*.xls; *.xlsx; *.et'), ('All Files', '*')])
        self.file_name = file_name
        print(file_name)
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
            self.set_button_state(0)

    # 使按钮失效，无法使用
    def set_button_state(self, i):
        if i == 0:
            self.open_file_button.config(state=DISABLED)
            self.export_file_button.config(state=DISABLED)
        elif i == 1:
            self.open_file_button.config(state=ACTIVE)
            self.export_file_button.config(state=ACTIVE)

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
        input_dialog = ExaminerDialog(name_list)
        self.set_button_state(0)
        self.init_window.wait_window(input_dialog.rootWindow)  # 这一句很重要！！！
        self.set_button_state(1)
        return input_dialog.result_list

    # 导出文件
    def export_file(self):
        if self.data is None:
            self.write_log("请先打开您需要处理的文件")
        else:
            res = self.open_export_dialog()
            print(res)
            if res is None:
                self.write_log("你取消了导出操作")
            else:
                self.proceed_data(res)

    def open_export_dialog(self):
        export_dialog = ExportDialog()
        self.set_button_state(0)
        self.init_window.wait_window(export_dialog.rootWindow)
        self.set_button_state(1)
        return export_dialog.result_list

    # 处理数据
    def proceed_data(self, res):
        if res[0] == 0:
            pass
        # 用户选择导出文件
        elif res[0] == 1:
            if res[1] == option1:
                pass
            if res[1] == option2:
                self.get_rate_all_solved()
                initial_filename = "员工事件成功解决率"
                # 开始导出
                file_name = filedialog.asksaveasfilename(title="保存文件",
                                                         filetype=[('表格文件', '*.xls')],
                                                         defaultextension='.xls',
                                                         initialfile=initial_filename)
                if file_name:
                    self.final_wb.save(file_name)
                    self.write_log('文件保存至：' + file_name)
                else:
                    self.write_log("文件名为空，导出中断。")
        # 用户选择导出图片
        elif res[0] == 2:
            pass

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

        ws = self.final_wb.add_sheet("员工根据解决率")
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

    # 添加日志
    def write_log(self, msg):  # 日志动态打印
        current_time = self.get_current_time()
        log_msg = str(current_time) + " " + str(msg) + "\n"  # 换行
        self.log_data_text.insert(END, log_msg)

    def exit_sys(self):
        self.init_window.destroy()
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


# 选择考核人员弹窗
class ExaminerDialog:
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


# 选择导出文件弹窗
class ExportDialog:
    def __init__(self):
        self.rootWindow = Toplevel()
        self.rootWindow.title('导出设置')
        self.rootWindow.geometry("300x180+250+250")
        self.result_list = []

        self.format_label = Label(self.rootWindow, text="导出格式")
        self.text_label = Label(self.rootWindow, text="导出内容")
        self.check_var1 = IntVar()
        self.check_var2 = IntVar()
        self.xls_cb = Checkbutton(self.rootWindow, text="导出文档", variable=self.check_var1, onvalue=1, offvalue=0,
                                  command=self.call_xls)
        self.img_cb = Checkbutton(self.rootWindow, text="导出图片", variable=self.check_var2, onvalue=1, offvalue=0,
                                  command=self.call_img)
        self.xls_cb.select()
        self.combo_var = StringVar()
        self.text_cb = ttk.Combobox(self.rootWindow, textvariable=self.combo_var)
        self.text_cb['values'] = (option1, option2, option3, option4, option5, option6, option7)
        self.text_cb['state'] = "readonly"
        self.text_cb.current(0)
        self.confirm_button = Button(self.rootWindow, text="确认", command=self.ok)
        self.cancel_button = Button(self.rootWindow, text="取消", command=self.cancel)
        self.init_ui()

    def call_xls(self):
        self.xls_cb.select()
        self.img_cb.deselect()

    def call_img(self):
        self.img_cb.select()
        self.xls_cb.deselect()

    def init_ui(self):
        self.format_label.place(relx=0.05, rely=0.05, relwidth=0.3, relheight=0.15)
        self.xls_cb.place(relx=0.05, rely=0.25, relwidth=0.3, relheight=0.1)
        self.img_cb.place(relx=0.6, rely=0.25, relwidth=0.3, relheight=0.1)
        self.text_label.place(relx=0.05, rely=0.4, relwidth=0.3, relheight=0.15)
        self.text_cb.place(relx=0.1, rely=0.60, relwidth=0.8, relheight=0.15)
        self.confirm_button.place(relx=0.6, rely=0.82, relwidth=0.15, relheight=0.15)
        self.cancel_button.place(relx=0.8, rely=0.82, relwidth=0.15, relheight=0.15)

    def ok(self):
        print(self.check_var1.get())
        if self.check_var1.get() == 1 and self.check_var2.get() == 0:
            self.result_list.append(1)
        elif self.check_var2.get() == 1 and self.check_var1.get() == 0:
            self.result_list.append(2)
        else:
            self.result_list.append(0)
        self.result_list.append(self.combo_var.get())
        self.rootWindow.destroy()

    def cancel(self):
        self.result_list = None  # 清空弹窗数据
        self.rootWindow.destroy()


# 数据类
class ExcelMaster:
    def __init__(self, data):
        self.data = data  # 源文件
        self.table = None  # 保存当前正在处理的表格
        # 初始化表格
        self.set_table(0)

    # index:第index个sheet,入参需要检查
    def set_table(self, index=0):
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
        print("正在查询: " + name)
        m = self.col_index('联系人')
        n = self.col_index('结束代码')
        name_list = list(self.table.col_values(m, start_rowx=1, end_rowx=None))
        code_list = list(self.table.col_values(n, start_rowx=1, end_rowx=None))
        print(name_list)
        print(code_list)
        print("length(name_list): " + str(len(name_list)))
        print("length(code_list): " + str(len(code_list)))
        # 遍历行
        solved_num = 0
        for i in range(len(name_list)):
            if name_list[i] == name and code_list[i] == '根本解决':
                solved_num += 1
        print("solved_num: " + str(solved_num))
        return solved_num

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
