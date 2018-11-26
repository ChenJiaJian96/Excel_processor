from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from xlrd import open_workbook, XLRDError
from xlwt import Workbook, Font, XFStyle
from time import strftime, localtime, mktime, strptime, time
from collections import Counter
import matplotlib.pyplot as plt
from matplotlib import rcParams

# TODO：导出文件加上工号
# 打包exe文件
# pyinstaller -F -w main.py

global option1, option2, option3, option4, option5, option6, option7
option1 = "全部导出"
option2 = "仅导出指定员工的“平均响应时长”情况"
option3 = "仅导出指定员工的“响应超时率”情况"
option4 = "仅导出指定员工的“按时解决率”情况"
option5 = "仅导出指定员工的“成功解决”情况"
option6 = "仅导出指定员工的“平均满意度”情况"
option7 = "仅导出指定员工的“平均解决时长”情况"

rcParams['font.sans-serif'] = ['SimHei']


# 主界面
class MyGUI:
    def __init__(self):
        self.file_name = None
        self.examiner_list = []  # 考核人员名单
        self.data = None  # master实例
        self.init_window = Tk()  # 父布局
        self.final_wb = None  # 最终导出的文件实例

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
        self.exclamation_label = Label(self.init_window, text=" ! ", font="bold, 8")
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
        self.log_data_text.place(relx=0.05, rely=0.2, relwidth=0.62, relheight=0.73)

        self.open_file_button.place(relx=0.72, rely=0.18, relwidth=0.18, relheight=0.17)
        self.export_file_button.place(relx=0.72, rely=0.38, relwidth=0.18, relheight=0.17)
        self.change_examiners_button.place(relx=0.72, rely=0.58, relwidth=0.18, relheight=0.17)
        self.exit_button.place(relx=0.72, rely=0.78, relwidth=0.18, relheight=0.17)

        self.log_scrollbar_y.config(command=self.log_data_text.yview)
        self.log_data_text.config(yscrollcommand=self.log_scrollbar_y.set)
        self.log_scrollbar_y.place(relx=0.67, rely=0.2, relheight=0.73)

        # 生成右侧提示按钮
        self.more_label.place(relx=0.93, rely=0.7, relwidth=0.03, relheight=0.08)
        self.question_label.place(relx=0.93, rely=0.8, relwidth=0.03, relheight=0.08)
        self.exclamation_label.place(relx=0.93, rely=0.9, relwidth=0.03, relheight=0.08)

        self.bottom_label.place(relx=0.4, rely=0.95, relwidth=0.2, relheight=0.05)

        self.set_button_state(0)

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
        if self.data.col_index('处理人') == -1:
            self.write_log("打开的文件中找不到列：“处理人”，无法导出员工名单")
            flag = 1
        if self.data.col_index('结束代码') == -1:
            self.write_log("打开的文件中找不到列：“结束代码”, 无法计算员工成功解决率")
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
            self.export_file_button.config(state=DISABLED)
            self.change_examiners_button.config(state=DISABLED)
        elif i == 1:
            self.export_file_button.config(state=ACTIVE)
            self.change_examiners_button.config(state=ACTIVE)

    # 设置员工列表
    def setup_staff_list(self):
        res = self.open_staff_list()
        print(res)
        if res is None or len(res) == 0:
            self.write_log("抱歉，你并未选择任何员工")
            if len(self.examiner_list) != 0:
                self.write_log("当前选择考核的员工为：" + str(self.examiner_list))
        else:
            self.write_log("当前选择考核的员工为：" + str(res))
            self.examiner_list = res
        # 若员工名单为空，不允许导出
        if len(self.examiner_list) == 0:
            self.export_file_button.config(state=DISABLED)

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
        # 首先清空缓存，避免重复导出之前的数据
        self.final_wb = Workbook(encoding='ascii')
        if len(res) == 0:
            self.write_log("你取消了导出")
        # 用户选择导出文件
        elif res[0] == 1:
            initial_filename = ""
            if res[1] == option1:
                initial_filename = "员工情况汇总表"
                pass
            if res[1] == option5:
                name_dict = self.get_rate_all_solved_data()
                self.get_rate_all_solved_xls(name_dict)
                initial_filename = "员工事件成功解决率"

            # 开始导出
            file_name = filedialog.asksaveasfilename(title="保存文件",
                                                     filetype=[('表格文件', '*.xls')],
                                                     defaultextension='.xls',
                                                     initialfile=initial_filename)
            try:
                self.final_wb.save(file_name)
            except PermissionError:
                self.write_log("权限出错，导出中断。")
            except FileNotFoundError:
                self.write_log("你点击了取消，导出中断。")
            else:
                self.write_log('导出成功。文件保存至：' + file_name)
        # 用户选择导出图片
        elif res[0] == 2:
            if res[1] == option5:
                name_dict = self.get_rate_all_solved_data()
                self.get_rate_all_solved_png(name_dict)
            else:
                pass

    # No.4:获取"事件成功解决率"的数据
    def get_rate_all_solved_data(self):
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
            cur_num = self.data.get_num_all_solved(name)  # 成功解决事件数
            rate = round(cur_num / total_num, 4)
            score = self.cal_score_all_solved(rate)
            name_dict[name] = [total_num, cur_num, rate * 100, score]
        # 输出结果
        for name in name_dict.keys():
            print(name + str(name_dict[name]))
        return name_dict

    # No.4:获取"事件成功解决率"文档
    def get_rate_all_solved_xls(self, name_dict):
        ws = self.final_wb.add_sheet("员工根据解决率")
        # 设置加粗字体
        style = XFStyle()
        font = Font()
        font.bold = True
        style.font = font
        ws.write(0, 0, "员工姓名", style=style)
        ws.write(0, 1, "事件完成数", style=style)
        ws.write(0, 2, "事件成功解决数", style=style)
        ws.write(0, 3, "事件成功解决率(%)", style=style)
        ws.write(0, 4, "该项得分", style=style)
        x = 1
        y = 0
        for name in name_dict.keys():
            ws.write(x, y, name, style=style)
            y = y + 1
            for item in name_dict[name]:
                ws.write(x, y, item)
                y = y + 1
            x = x + 1
            y = 0

    # No.4:获取"事件成功解决率"导出图片
    def get_rate_all_solved_png(self, name_dict):
        label_list, num_list1, num_list2, num_list3, num_list4 = [], [], [], [], []
        for name in name_dict:
            label_list.append(name)
            num_list1.append(name_dict[name][0])
            num_list2.append(name_dict[name][1])
            num_list3.append(name_dict[name][2])
            num_list4.append(name_dict[name][3])
        x = range(len(num_list1))
        # 设置画布大小
        plt.figure(figsize=(len(label_list) + 7, 5))
        rects1 = plt.bar(x=[i + 0.2 for i in x], height=num_list1, width=0.2, alpha=0.8, color='red', label="事件完成数")
        rects2 = plt.bar(x=[i + 0.4 for i in x], height=num_list2, width=0.2, color='green', label="事件成功解决数")
        rects3 = plt.bar(x=[i + 0.6 for i in x], height=num_list3, width=0.2, color='blue', label="事件成功解决率(%)")
        rects4 = plt.bar(x=[i + 0.8 for i in x], height=num_list4, width=0.2, color='yellow', label="该项得分")
        # 取值范围
        plt.ylim(0, max(max(num_list1), 105))
        plt.xlim(0, len(label_list))
        # 中点坐标，显示值
        plt.xticks([index + 0.5 for index in x], label_list)
        plt.xlabel("员工姓名")
        plt.ylabel("数量（得分）")
        plt.title("员工成功解决率统计图表")
        plt.legend(bbox_to_anchor=(1.01, 1), loc=2, borderaxespad=0., handleheight=1.675)
        #  编辑文本
        for rect in rects1:
            height = rect.get_height()
            plt.text(rect.get_x() + rect.get_width() / 2, height + 1, str(height), ha="center", va="bottom")
        for rect in rects2:
            height = rect.get_height()
            plt.text(rect.get_x() + rect.get_width() / 2, height + 1, str(height), ha="center", va="bottom")
        for rect in rects3:
            height = rect.get_height()
            plt.text(rect.get_x() + rect.get_width() / 2, height + 1, str(height), ha="center", va="bottom")
        for rect in rects4:
            height = rect.get_height()
            plt.text(rect.get_x() + rect.get_width() / 2, height + 1, str(height), ha="center", va="bottom")
        plt.tight_layout()
        initial_filename = "员工成功解决率统计图表"
        filename = filedialog.asksaveasfilename(title="保存文件",
                                                filetype=[('图片文件', '*.png')],
                                                defaultextension='.png',
                                                initialfile=initial_filename)
        try:
            plt.savefig(filename)
        except PermissionError:
            self.write_log("权限出错，导出中断。")
        except FileNotFoundError:
            self.write_log("你点击了取消，导出中断。")
        else:
            self.write_log("导出图表成功，文件保存至：" + filename)
        plt.show()

    # 添加日志
    def write_log(self, msg):  # 日志动态打印
        current_time = self.get_current_time()
        log_msg = str(current_time) + " " + str(msg) + "\n"  # 换行
        self.log_data_text.insert(END, log_msg)
        divider_msg = "---------------------------------------" + "\n"
        self.log_data_text.insert(END, divider_msg)
        # 滚动至底部
        self.log_data_text.yview_moveto(100)

    def exit_sys(self):
        self.init_window.destroy()
        quit()

    @staticmethod
    def get_current_time():
        current_time = strftime('%Y-%m-%d %H:%M:%S', localtime(time()))
        return current_time

    # 计算事件平均响应时长的得分
    @staticmethod
    def cal_score_ave_response(hour):
        if hour <= 0.2:
            return 100
        elif hour <= 0.5:
            return 90
        elif hour <= 1:
            return 80
        elif hour <= 3:
            return 70
        else:
            return 60

    # 计算事件响应超时率的得分
    @staticmethod
    def cal_score_overtime(rate):
        if rate <= 0.001:
            return 100
        elif rate <= 0.01:
            return 90
        elif rate <= 0.1:
            return 80
        elif rate <= 0.2:
            return 70
        else:
            return 60

    # 计算事件按时解决率的得分
    @staticmethod
    def cal_score_on_time(rate):
        if rate >= 0.997:
            return 100
        elif rate >= 0.99:
            return 90
        elif rate >= 0.9:
            return 80
        elif rate >= 0.8:
            return 70
        else:
            return 60

    # 计算成功解决率的得分
    @staticmethod
    def cal_score_all_solved(rate):
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

    # 计算用户平均满意度的得分
    @staticmethod
    def cal_score_ave_satisfied(level):
        if level >= 99.5:
            return 100
        elif level >= 98:
            return 90
        elif level >= 80:
            return 80
        elif level >= 70:
            return 70
        else:
            return 10 * int(level / 10)

    # 计算工作能力得分
    @staticmethod
    def cal_score_work_ability(hour):
        if hour <= 1:
            return 100
        elif hour <= 4:
            return 90
        elif hour <= 12:
            return 80
        elif hour <= 48:
            return 70
        else:
            return 60


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

        # 滚动条
        self.box_scrollbar_y = Scrollbar(self.rootWindow)

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
        self.box_scrollbar_y.config(command=self.name_list_box.yview)
        self.name_list_box.config(yscrollcommand=self.box_scrollbar_y.set)
        self.box_scrollbar_y.place(relx=0.35, rely=0.3, relheight=0.65)
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
        self.value_list1 = (option1, option2, option3, option4, option5, option6, option7)
        self.value_list2 = (option2, option3, option4, option5, option6, option7)
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
        self.img_cb = Checkbutton(self.rootWindow, text="仅导出图片", variable=self.check_var2, onvalue=1, offvalue=0,
                                  command=self.call_img)
        self.xls_cb.select()
        self.combo_var = StringVar()
        self.text_cb = ttk.Combobox(self.rootWindow, textvariable=self.combo_var)
        self.text_cb['values'] = self.value_list1
        self.text_cb['state'] = "readonly"
        self.text_cb.current(0)
        self.confirm_button = Button(self.rootWindow, text="确认", command=self.ok)
        self.cancel_button = Button(self.rootWindow, text="取消", command=self.cancel)
        self.init_ui()

    def call_xls(self):
        self.xls_cb.select()
        self.img_cb.deselect()
        self.text_cb['values'] = self.value_list1
        self.text_cb.current(0)

    def call_img(self):
        self.img_cb.select()
        self.xls_cb.deselect()
        self.text_cb['values'] = self.value_list2
        self.text_cb.current(0)

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
        i = self.col_index('处理人')
        name_dict = Counter(self.table.col_values(i, start_rowx=1, end_rowx=None))
        return list(name_dict.keys())

    # 返回表格的员工完成事件数
    def get_name_dict(self):
        i = self.col_index('处理人')
        name_dict = Counter(self.table.col_values(i, start_rowx=1, end_rowx=None))
        return name_dict

    # 返回员工“成功解决”的事件总数
    def get_num_all_solved(self, name):
        print("正在查询: " + name)
        m = self.col_index('处理人')
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
