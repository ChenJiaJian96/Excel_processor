# -*- coding: utf-8 -*-
from tkinter import *
from tkinter import filedialog, messagebox, scrolledtext, ttk
from xlrd import open_workbook, XLRDError, xldate_as_tuple
from xlwt import Workbook, Font, XFStyle
from time import strftime, localtime, mktime, strptime, time
from datetime import datetime
from collections import Counter
import matplotlib.pyplot as plt
from matplotlib import rcParams
from decimal import Decimal

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
global ico_path
ico_path = ".\CSPGCL.ico"
global color_scheme
color_scheme = [['#5B6C83', '#D7CCB8', '#38526E', '#BFBFBF'], ['#948A54', '#596166', "#A9BD8B", "#1C7B64"],
                ['#1A7F9C', '#2DCFFF', '#104D60', "#229BBF"]]
rcParams['font.sans-serif'] = ['SimHei']


# 主界面
class MyGUI:
    def __init__(self):
        self.file_name = None
        self.examiner_list = []  # 考核人员名单
        self.history_list = []  # 历史考核人员名单
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
        self.init_window.title("IT服务精益化管理系统")  # 指定标题
        self.init_window.geometry("500x265+100+100")  # 指定初始化大小以及出现位置
        # self.init_window.attributes("-alpha", 0.8)  # 指定透明度
        self.init_window.iconbitmap(ico_path)

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
        self.more_label.bind("<Button-1>", self.show_score_standard)
        self.more_label.place(relx=0.93, rely=0.7, relwidth=0.03, relheight=0.08)
        self.question_label.bind("<Button-1>", self.show_instruction)
        self.question_label.place(relx=0.93, rely=0.8, relwidth=0.03, relheight=0.08)
        self.exclamation_label.bind("<Button-1>", func=self.show_software_detail)
        self.exclamation_label.place(relx=0.93, rely=0.9, relwidth=0.03, relheight=0.08)

        self.bottom_label.place(relx=0.4, rely=0.95, relwidth=0.2, relheight=0.05)

        self.set_button_state(0)
        _time = localtime(time())
        greetings = self.get_greetings(_time.tm_hour)
        self.write_log(greetings + ",请点击右侧“打开文件”按钮开始本次考核吧-->")

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
            self.write_log("打开的文件中找不到列：“处理人”")
            flag = 1
        if self.data.col_index('结束代码') == -1:
            self.write_log("打开的文件中找不到列：“结束代码”")
            flag = 1
        if self.data.col_index('派单时间') == -1:
            self.write_log("打开的文件中找不到列：“派单时间”")
            flag = 1
        if self.data.col_index('完成时间') == -1:
            self.write_log("打开的文件中找不到列：“完成时间”")
            flag = 1
        if self.data.col_index('销单时间') == -1:
            self.write_log("打开的文件中找不到列：“销单时间”")
            flag = 1
        if self.data.col_index('处理时间(小时)') == -1:
            self.write_log("打开的文件中找不到列：“处理时间(小时)”")
            flag = 1
        if self.data.col_index('事件优先级') == -1:
            self.write_log("打开的文件中找不到列：“事件优先级”")
            flag = 1
        if flag == 0:
            self.write_log("该文件完整，开始选择考勤名单。")
            self.history_list = []
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

    # 显示评分标准
    def show_score_standard(self, event):
        standard_dialog = StandardDialog()
        self.init_window.wait_window(standard_dialog.rootWindow)

    # 显示软件详情
    @staticmethod
    def show_software_detail(event):
        messagebox.showinfo("关于", "ISBN:\n著作权人:\n出版单位:")

    # 显示操作说明
    def show_instruction(self, event):
        instruction_dialog = InstructionDialog()
        self.init_window.wait_window(instruction_dialog.rootWindow)

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
            if str(res) not in self.history_list:
                self.history_list.append(res)
            self.examiner_list = res
        # 若员工名单为空，不允许导出
        if len(self.examiner_list) == 0:
            self.export_file_button.config(state=DISABLED)

    # 开始选择员工列表
    def open_staff_list(self):
        name_list = self.data.get_name_list()
        input_dialog = ExaminerDialog(name_list, self.history_list)
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

    # 打开导出文件弹窗
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
                name_dict = self.get_ave_response_data()
                self.get_ave_response_xls(name_dict)
                name_dict = self.get_over_time_data()
                self.get_over_time_xls(name_dict)
                name_dict = self.get_on_time_data()
                self.get_on_time_xls(name_dict)
                name_dict = self.get_rate_all_solved_data()
                self.get_rate_all_solved_xls(name_dict)
                name_dict = self.get_rate_ave_satisfied_data()
                self.get_rate_ave_satisfied_xls(name_dict)
                name_dict = self.get_ave_solved_data()
                self.get_ave_solved_xls(name_dict)
                initial_filename = "员工情况汇总表"
            elif res[1] == option2:
                name_dict = self.get_ave_response_data()
                self.get_ave_response_xls(name_dict)
                initial_filename = "事件平均响应时长"
            elif res[1] == option3:
                name_dict = self.get_over_time_data()
                self.get_over_time_xls(name_dict)
                initial_filename = "事件响应超时率"
            elif res[1] == option4:
                name_dict = self.get_on_time_data()
                self.get_on_time_xls(name_dict)
                initial_filename = "事件按时解决率"
            elif res[1] == option5:
                name_dict = self.get_rate_all_solved_data()
                self.get_rate_all_solved_xls(name_dict)
                initial_filename = "员工事件成功解决率"
            elif res[1] == option6:
                name_dict = self.get_rate_ave_satisfied_data()
                self.get_rate_ave_satisfied_xls(name_dict)
                initial_filename = "客户平均满意度"
            elif res[1] == option7:
                name_dict = self.get_ave_solved_data()
                self.get_ave_solved_xls(name_dict)
                initial_filename = "事件平均解决时长"

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
            if res[1] == option1:
                name_dict_list = []
                name_list = self.examiner_list
                name_dict_list.append(self.get_ave_response_data())
                name_dict_list.append(self.get_over_time_data())
                name_dict_list.append(self.get_on_time_data())
                name_dict_list.append(self.get_rate_all_solved_data())
                name_dict_list.append(self.get_rate_ave_satisfied_data())
                name_dict_list.append(self.get_ave_solved_data())
                self.get_total_png_by_data(name_list, name_dict_list)
            elif res[1] == option2:
                name_dict = self.get_ave_response_data()
                png_element = ["事件完成数", "事件响应总时长(h)", "事件平均响应时长(h)", "该项得分", "处理人", "数量（时长/得分）", "员工平均响应时长统计图表",
                               "员工平均响应时长统计图表"]
                self.get_png_by_data(name_dict, png_element, color_scheme[1])
                pass
            elif res[1] == option3:
                name_dict = self.get_on_time_data()
                png_element = ["事件完成数", "事件超时解决数", "事件超时解决率(%)", "该项得分", "处理人", "数量（得分）", "员工超时解决率统计图表", "员工超时解决率统计图表"]
                self.get_png_by_data(name_dict, png_element, color_scheme[2])
            elif res[1] == option4:
                name_dict = self.get_on_time_data()
                png_element = ["事件完成数", "事件按时解决数", "事件按时解决率(%)", "该项得分", "处理人", "数量（得分）", "员工按时解决率统计图表", "员工超时解决率统计图表"]
                self.get_png_by_data(name_dict, png_element, color_scheme[2])
            elif res[1] == option5:
                name_dict = self.get_rate_all_solved_data()
                png_element = ["事件完成数", "事件成功解决数", "事件成功解决率(%)", "该项得分", "处理人", "数量（得分）", "员工成功解决率统计图表", "员工成功解决率统计图表"]
                self.get_png_by_data(name_dict, png_element, color_scheme[2])
            elif res[1] == option6:
                name_dict = self.get_rate_ave_satisfied_data()
                png_element = ["事件完成数", "客户满意数", "事件满意率(%)", "该项得分", "处理人", "数量（得分）", "客户平均满意率统计图表", "员工成功解决率统计图表"]
                self.get_png_by_data(name_dict, png_element, color_scheme[2])
            elif res[1] == option7:
                name_dict = self.get_ave_solved_data()
                png_element = ["事件完成数", "事件解决总时长(h)", "事件平均解决时长(h)", "该项得分", "处理人", "数量（时长/得分）", "员工平均解决时长统计图表",
                               "员工平均解决时长统计图表"]
                self.get_png_by_data(name_dict, png_element, color_scheme[0])
                pass

    # No.1:获取"事件平均响应时长"的数据
    def get_ave_response_data(self):
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
            sum_res_h = Decimal(self.data.get_total_response(name) / 3600).quantize(Decimal('0.0000'))  # 总响应时间(h)
            ave_res_h = Decimal(sum_res_h / total_num).quantize(Decimal('0.00'))
            score = self.cal_score_ave_response(ave_res_h)
            name_dict[name] = [total_num, sum_res_h, ave_res_h, score]
        return name_dict

    # No.1:获取"事件平均响应时长"的文档
    def get_ave_response_xls(self, name_dict):
        ws = self.final_wb.add_sheet("事件平均响应时长")
        # 设置加粗字体
        style = XFStyle()
        font = Font()
        font.bold = True
        style.font = font
        ws.write(0, 0, "员工姓名", style=style)
        ws.write(0, 1, "事件完成数", style=style)
        ws.write(0, 2, "事件总响应时间(h)", style=style)
        ws.write(0, 3, "事件平均响应时间(h)", style=style)
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

    # No.2:获取"事件超时解决率"的数据
    def get_over_time_data(self):
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
            cur_num = total_num - self.data.get_num_solved_ontime(name)  # 超时解决事件数
            rate = Decimal(cur_num / total_num).quantize(Decimal('0.0000'))
            score = self.cal_score_overtime(rate)
            name_dict[name] = [total_num, cur_num, rate * 100, score]
        return name_dict

    # No.2:获取"事件超时解决率"的文档
    def get_over_time_xls(self, name_dict):
        ws = self.final_wb.add_sheet("超时解决率")
        # 设置加粗字体
        style = XFStyle()
        font = Font()
        font.bold = True
        style.font = font
        ws.write(0, 0, "员工姓名", style=style)
        ws.write(0, 1, "事件完成数", style=style)
        ws.write(0, 2, "事件超时解决数", style=style)
        ws.write(0, 3, "事件超时解决率(%)", style=style)
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

    # No.3:获取"事件按时解决率"的数据
    def get_on_time_data(self):
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
            cur_num = self.data.get_num_solved_ontime(name)  # 按时解决事件数
            rate = Decimal(cur_num / total_num).quantize(Decimal('0.0000'))
            score = self.cal_score_on_time(rate)
            name_dict[name] = [total_num, cur_num, rate * 100, score]
        return name_dict

    # No.3:获取"事件按时解决率"的文档
    def get_on_time_xls(self, name_dict):
        ws = self.final_wb.add_sheet("按时解决率")
        # 设置加粗字体
        style = XFStyle()
        font = Font()
        font.bold = True
        style.font = font
        ws.write(0, 0, "员工姓名", style=style)
        ws.write(0, 1, "事件完成数", style=style)
        ws.write(0, 2, "事件按时解决数", style=style)
        ws.write(0, 3, "事件按时解决率(%)", style=style)
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
            rate = Decimal(cur_num / total_num).quantize(Decimal('0.000'))
            score = self.cal_score_all_solved(rate)
            name_dict[name] = [total_num, cur_num, rate * 100, score]
        return name_dict

    # No.4:获取"事件成功解决率"文档
    def get_rate_all_solved_xls(self, name_dict):
        ws = self.final_wb.add_sheet("根本解决率")
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

    # 根据数据导出全部图片
    def get_total_png_by_data(self, name_list, name_dict_list):
        num_list_list = []
        for name in name_list:
            num_list = []
            for cur_dict in name_dict_list:
                num_list.append(cur_dict[name][3])
            num_list_list.append(num_list)

        label_list = ["事件平均解决时长", "事件响应超时率", "事件按时解决率", "事件成功解决率", "客户平均满意度", "事件平均解决时长"]
        x = range(len(label_list))
        # 设置画布大小
        plt.figure(figsize=(len(label_list) + 7, 5))
        rect_list = []
        rect_width = 1 / (len(name_list) + 1)
        for m in range(len(name_list)):
            rects = plt.bar(x=[i + rect_width * (m + 1) for i in x], height=num_list_list[m], width=rect_width,
                            color='#4F94CD', edgecolor='k', label=name_list[m])
            rect_list.append(rects)
        # 取值范围
        plt.ylim(0, 105)
        plt.xlim(0, len(label_list))
        # 中点坐标，显示值
        plt.xticks([index + 0.5 for index in x], label_list)
        plt.xlabel("评判标准")
        plt.ylabel("该项得分")
        plt.title("员工各项标准得分")
        plt.legend(bbox_to_anchor=(1.01, 1), loc=2, borderaxespad=0., handleheight=1.675)
        #  编辑文本
        for cur_rect in rect_list:
            for rect in cur_rect:
                height = rect.get_height()
                plt.text(rect.get_x() + rect.get_width() / 2, height + 1, str(height), ha="center", va="bottom")
        plt.tight_layout()
        initial_filename = "员工各项标准得分情况统计表"
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

    # 根据数据导出个别图片
    def get_png_by_data(self, name_dict, png_element, color_scheme):
        label_list, num_list1, num_list2, num_list3, num_list4 = [], [], [], [], []
        for name in name_dict:
            label_list.append(name)
            num_list1.append(name_dict[name][0])
            num_list2.append(name_dict[name][1])
            num_list3.append(name_dict[name][2])
            num_list4.append(name_dict[name][3])
        x = range(len(label_list))
        # 设置画布大小
        plt.figure(figsize=(len(label_list) + 7, 5))
        rects1 = plt.bar(x=[i + 0.2 for i in x], height=num_list1, width=0.2, color=color_scheme[0], edgecolor='k',
                         label=png_element[0])
        rects2 = plt.bar(x=[i + 0.4 for i in x], height=num_list2, width=0.2, color=color_scheme[1], edgecolor='k',
                         label=png_element[1])
        rects3 = plt.bar(x=[i + 0.6 for i in x], height=num_list3, width=0.2, color=color_scheme[2], edgecolor='k',
                         label=png_element[2])
        rects4 = plt.bar(x=[i + 0.8 for i in x], height=num_list4, width=0.2, color=color_scheme[3], edgecolor='k',
                         label=png_element[3])
        # 取值范围
        plt.ylim(0, max(max(num_list1), 100) * 1.2)
        plt.xlim(0, len(label_list))
        # 中点坐标，显示值
        plt.xticks([index + 0.5 for index in x], label_list)
        plt.xlabel(png_element[4])
        plt.ylabel(png_element[5])
        plt.title(png_element[6])
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
        plt.show()
        # initial_filename = png_element[7]
        # filename = filedialog.asksaveasfilename(title="保存文件",
        #                                         filetype=[('图片文件', '*.png')],
        #                                         defaultextension='.png',
        #                                         initialfile=initial_filename)
        # try:
        #     plt.savefig(filename)
        # except PermissionError:
        #     self.write_log("权限出错，导出中断。")
        # except FileNotFoundError:
        #     self.write_log("你点击了取消，导出中断。")
        # else:
        #     self.write_log("导出图表成功，文件保存至：" + filename)

    # No.5:获取"客户平均满意度"数据
    def get_rate_ave_satisfied_data(self):
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
            s_sum = total_num
            rate = 1
            score = 100
            name_dict[name] = [total_num, s_sum, rate * 100, score]
        return name_dict

    # No.5:获取"客户平均满意度"文档
    def get_rate_ave_satisfied_xls(self, name_dict):
        ws = self.final_wb.add_sheet("客户平均满意度")
        # 设置加粗字体
        style = XFStyle()
        font = Font()
        font.bold = True
        style.font = font
        ws.write(0, 0, "员工姓名", style=style)
        ws.write(0, 1, "事件完成数", style=style)
        ws.write(0, 2, "客户满意数", style=style)
        ws.write(0, 3, "事件满意率(%)", style=style)
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

    # No.6:获取"事件平均解决时长"数据
    def get_ave_solved_data(self):
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
            solved_time = Decimal(self.data.get_total_solved_time(name)).quantize(Decimal('0.000'))
            ave_solved_time = Decimal(solved_time / total_num).quantize(Decimal('0.000'))
            score = self.cal_score_ave_solved(ave_solved_time)
            name_dict[name] = [total_num, solved_time, ave_solved_time, score]
        return name_dict

    # No.6:获取"事件平均解决时长"文档
    def get_ave_solved_xls(self, name_dict):
        ws = self.final_wb.add_sheet("事件平均解决时长")
        # 设置加粗字体
        style = XFStyle()
        font = Font()
        font.bold = True
        style.font = font
        ws.write(0, 0, "员工姓名", style=style)
        ws.write(0, 1, "事件完成数", style=style)
        ws.write(0, 2, "事件解决总时长(h)", style=style)
        ws.write(0, 3, "事件平均解决时长(h)", style=style)
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

    # 添加日志
    def write_log(self, msg):  # 日志动态打印
        current_time = self.get_current_time()
        log_msg = str(current_time) + " " + str(msg) + "\n"  # 换行
        self.log_data_text.insert(END, log_msg)
        divider_msg = "-----------------------------------------" + "\n"
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

    @staticmethod
    def get_greetings(hour):
        if 6 <= hour <= 11:
            return "早上好"
        elif 11 <= hour <= 13:
            return "中午好"
        elif 13 <= hour <= 18:
            return "下午好"
        else:
            return "晚上好"

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
            return 60

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
            return 60

    # 计算事件平均解决时长的得分
    @staticmethod
    def cal_score_ave_solved(hour):
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
    def __init__(self, name_list, history_list):
        self.name_list = name_list  # 传过来的名单
        self.history_list = history_list  # 传过来的历史选择名单
        self.result_list = []  # 需要发出去的名单
        self.rootWindow = Toplevel()
        self.rootWindow.title('设置考勤名单')
        self.rootWindow.geometry("600x380+250+250")
        self.rootWindow.iconbitmap(ico_path)
        self.search_text = Entry(self.rootWindow)
        self.name_list_label = Label(self.rootWindow, text="表格名单(点击多选）")
        self.selected_list_label = Label(self.rootWindow, text="选中名单列表")
        self.history_label = Label(self.rootWindow, text="历史选择")
        self.history_var = StringVar()
        self.history_cb = ttk.Combobox(self.rootWindow, textvariable=self.history_var)
        self.history_cb['values'] = self.history_list
        self.history_cb['state'] = "readonly"

        self.search_button = Button(self.rootWindow, text="搜索", command=self.search)
        self.add_button = Button(self.rootWindow, text="添加 >>", command=self.add_name_from_box)
        self.del_button = Button(self.rootWindow, text="删除 <<", command=self.del_name)
        self.all_del_button = Button(self.rootWindow, text="全部删除", command=self.del_all)
        self.confirm_button = Button(self.rootWindow, text="确认", command=self.ok)
        self.history_add_button = Button(self.rootWindow, text="↑↑", command=self.add_name_from_cb)
        # 滚动条
        self.box_scrollbar_y = Scrollbar(self.rootWindow)

        self.name_list_box = Listbox(self.rootWindow, selectmode=MULTIPLE)  # 表格员工名单
        self.selected_list_box = Listbox(self.rootWindow, selectmode=BROWSE)  # 选中员工名单
        # 弹窗界面
        self.init_ui()

    def init_ui(self):
        self.search_text.place(relx=0.05, rely=0.05, relwidth=0.6, relheight=0.08)
        self.search_button.place(relx=0.7, rely=0.05, relwidth=0.25, relheight=0.08)
        self.name_list_label.place(relx=0.05, rely=0.14, relwidth=0.3, relheight=0.1)
        self.selected_list_label.place(relx=0.65, rely=0.14, relwidth=0.3, relheight=0.1)
        self.name_list_box.place(relx=0.05, rely=0.24, relwidth=0.3, relheight=0.65)
        self.selected_list_box.place(relx=0.65, rely=0.24, relwidth=0.3, relheight=0.65)
        self.add_button.place(relx=0.4, rely=0.24, relwidth=0.2, relheight=0.12)
        self.del_button.place(relx=0.4, rely=0.41, relwidth=0.2, relheight=0.12)
        self.all_del_button.place(relx=0.4, rely=0.58, relwidth=0.2, relheight=0.12)
        self.confirm_button.place(relx=0.4, rely=0.75, relwidth=0.2, relheight=0.12)
        self.box_scrollbar_y.config(command=self.name_list_box.yview)
        self.name_list_box.config(yscrollcommand=self.box_scrollbar_y.set)
        self.box_scrollbar_y.place(relx=0.35, rely=0.24, relheight=0.65)
        self.history_label.place(relx=0.05, rely=0.9, relwidth=0.10, relheight=0.08)
        self.history_cb.place(relx=0.2, rely=0.9, relwidth=0.6, relheight=0.08)
        self.history_add_button.place(relx=0.85, rely=0.9, relwidth=0.1, relheight=0.08)
        if len(self.history_list) == 0:
            self.history_add_button.config(state=DISABLED)
        # 对名单进行排序，优化用户体验
        self.refresh_name_list()

    def refresh_name_list(self):
        self.name_list_box.delete(0, END)
        try:
            engine = SortEngine()
            self.name_list = engine.cnsort(self.name_list)
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

    def add_name(self, name_list):
        temp_added_list = []
        for pos in name_list:
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

    def add_name_from_box(self):
        selected_list = self.name_list_box.curselection()
        print(selected_list)
        self.add_name(selected_list)

    def add_name_from_cb(self):
        self.del_all()
        temp_name_list = self.history_list[self.history_cb.current()]
        print(self.name_list)
        print(temp_name_list)
        selected_list = []
        for name in temp_name_list:
            print(name)
            pos = self.name_list.index(name)
            selected_list.append(pos)
        self.add_name(selected_list)

    def ok(self):
        self.rootWindow.destroy()

    def cancel(self):
        self.result_list = None  # 清空弹窗数据
        self.rootWindow.destroy()

    # 模糊搜索
    @staticmethod
    def fuzzyfinder(user_input, collection):
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
        self.value_list = (option1, option2, option3, option4, option5, option6, option7)
        self.rootWindow = Toplevel()
        self.rootWindow.title('导出设置')
        self.rootWindow.geometry("300x180+250+250")
        self.rootWindow.iconbitmap(ico_path)
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
        self.text_cb['values'] = self.value_list
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


# 显示评分标准弹窗
class StandardDialog:
    def __init__(self):
        self.rootWindow = Toplevel()
        self.rootWindow.title('评分标准')
        self.rootWindow.geometry("780x580+250+250")
        self.rootWindow.iconbitmap(ico_path)
        style = ttk.Style()
        style.configure('Calendar.Treeview', rowheight=90)
        tree = ttk.Treeview(self.rootWindow, show="headings", style='Calendar.Treeview')  # 表格
        tree["columns"] = ("序号", "项目", "单位", "指标", "描述", "评分标准")
        tree.column("序号", width=40, anchor=CENTER)
        tree.column("项目", width=110, anchor=CENTER)
        tree.column("单位", width=40, anchor=CENTER)
        tree.column("指标", width=80)
        tree.column("描述", width=330)
        tree.column("评分标准", width=180)

        tree.heading("序号", text="序号")
        tree.heading("项目", text="项目")
        tree.heading("单位", text="单位")
        tree.heading("指标", text="指标")
        tree.heading("描述", text="描述")
        tree.heading("评分标准", text="评分标准")

        tree.insert("", 0, values=(
            "1", "事件平均响应时长", "小时", "地市局指标", "已响应事件工单响应总时长/已响应事件工单总数*100%。\n反映每张事件工单的平均响应时长。",
            "<=0.2, 100分；\n<=0.5且>0.2, 90分；\n<=1且>0.5, 80分；\n<=3且>1, 70分；\n>3, 60分及以下\n"), tags='T')
        tree.insert("", 1, values=(
            "2", "事件响应超时率", "%", "地市局指标", "响应超时事件工单总量/事件工单总量*100%。\n反映1000号运维人员对事件工单进行响应的超时情况。",
            "<=0.1%, 100分；\n<=1%且>0.1%, 90分；\n<=10%且>1%, 80分；\n<=20%且>10%, 70分；\n>20%, 60分及以下\n"), tags='T')
        tree.insert("", 2, values=(
            "3", "事件按时解决率", "%", "地市局指标", "按时解决事件工单总量/事件工单总量*100%。\n反映1000号运维人员是否能够在服务等级（SLA）协议内\n解决事件工单的情况。",
            ">=99.7%, 100分；\n>=99%且<99.7%, 90分；\n>=90%且<99%, 80分；\n>=80%且<90%, 70分；\n<80%, 60分及以下\n"), tags='T')
        tree.insert("", 3, values=(
            "4", "事件成功解决率", "%", "地市局指标", "事件关闭代码为“根本解决”的事件总数/事件总数*100%。\n反映事件根本解决的能力。",
            ">=99.5%, 100分；\n>=98%且<99.5%, 90分；\n>=80%且<98%, 80分；\n>=70%且<80%, 70分；\n<70%, 60分及以下\n"), tags='T')
        tree.insert("", 4, values=("5", "客户平均满意度", "", "地市局指标", "Σ满意度/统计次数\n反映客户对IT服务的程度情况",
                                   ">=99.5%, 100分；\n>=98%且<99.5%, 90分；\n>=80%且<98%, 80分；\n>=70%且<80%, 70分；\n<70%, "
                                   "60分及以下\n"), tags='T')
        tree.insert("", 5, values=(
            "6", "事件平均解决时长", "小时", "地市局指标", "已解决事件工单解决总时长/已解决事件工单总数*100%。\n反映每张事件工单的平均解决时长",
            "<=1, 100分；\n<=4且>1, 90分；\n<=12且>4, 80分；\n<=48且>12, 70分；\n>48, 60分及以下\n"), tags='T')
        tree.tag_configure('T')
        tree.place(relx=0, rely=0, relwidth=1, relheight=1)


# 显示使用流程弹窗
class InstructionDialog:
    def __init__(self):
        self.rootWindow = Toplevel()
        self.rootWindow.title('使用流程和常见问题')
        self.rootWindow.geometry("500x400+250+250")
        self.rootWindow.iconbitmap(ico_path)

        self.guide_button = Button(self.rootWindow, text="使用流程", command=lambda: self.update_text(1))
        self.quest_button = Button(self.rootWindow, text="常见问题", command=lambda: self.update_text(2))
        self.wel_text = "欢迎查阅使用流程及常见问题\n\n请点击上面按钮进行查询↑↑↑"
        self.guide_text = "使用说明\n\n" \
                          "一、使用流程\n" \
                          "打开文件->修改考核人员名单->导出文件->退出系统\n\n" \
                          "二、打开文件\n" \
                          "2.1--本系统可以打开常用的表格文件，如.xlsx/.xls/.et等文件。请在打开文件窗中选中需要打开的文件，" \
                          "若文件过大需要耐心等待一段时间，未打开文件将无法使用导出/修改考核名单功能；\n" \
                          "2.2--文件完整性检查，为了使本系统的功能均能正常使用，打开文件后会进行完整性检查，主要检查文件中是否" \
                          "存在以下列，包括‘处理人’、‘结束代码’等，若数据不完整，将无法进行进一步导出。\n\n" \
                          "三、修改考核人员名单\n" \
                          "3.1--鉴于导入文件可能数据量过大，并针对软件开发需求，导出数据将从‘处理人’列中的个别或全部项计算得出，" \
                          "导出文件会按照选择的考核名单进行针对性导出；\n" \
                          "3.2--在成功导入完整数据文件后，会直接进入修改考核人员名单的界面，而在导出文件前的任意时刻也可以选择更改，" \
                          "若使用者在选择界面中未选择任何人员，会默认保持为最近一次的名单；\n" \
                          "3.3--修改考核人员名单界面中，用户可以在左侧的名单列表中选择需要添加的名单，点击添加按钮即可添加，" \
                          "添加成功后的人员会显示在右侧栏，相应地，可以选择删除对应人员名单；\n" \
                          "3.4--上述界面搜索功能，用户可以在位于界面上侧搜索栏搜索对应人员，并执行后续操作，当搜索栏为空并选择‘搜索’按钮后，" \
                          "名单会从搜索特定名单恢复为全部名单；\n" \
                          "3.5--完成选择后点击确认按钮返回主界面，并完成考核名单的修改。\n\n" \
                          "四、导出文件\n" \
                          "4.1--在导出文件弹窗中balabala"
        self.quest_text = "常见问题说明\n\n" \
                          "1.为什么选择考核名单弹窗->表格名单中"
        self.content_text = scrolledtext.ScrolledText(self.rootWindow, wrap=WORD)
        self.box_scrollbar_y = Scrollbar(self.rootWindow)

        self.guide_button.place(relx=0.27, rely=0.03, relwidth=0.2, relheight=0.1)
        self.quest_button.place(relx=0.53, rely=0.03, relwidth=0.2, relheight=0.1)
        self.content_text.place(relx=0.02, rely=0.16, relwidth=0.96, relheight=0.81)
        self.update_text(0)

    def update_text(self, update_type):
        self.content_text.delete(1.0, END)
        if update_type == 1:
            self.content_text.insert(INSERT, self.guide_text)
        elif update_type == 2:
            self.content_text.insert(INSERT, self.quest_text)
        else:
            self.content_text.insert(INSERT, self.wel_text)


# 数据类
class ExcelMaster:
    def __init__(self, data):
        self.data = data  # 源文件
        self.table = None  # 保存当前正在处理的表格
        # 初始化表格
        self.set_table(0)
        # 获取表格总行数
        self.nrow = self.table.nrows
        print(self.nrow)

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
        if name_dict.keys().__contains__(""):
            del name_dict[""]
        return list(name_dict.keys())

    # 返回表格的员工完成事件数
    def get_name_dict(self):
        i = self.col_index('处理人')
        name_dict = Counter(self.table.col_values(i, start_rowx=1, end_rowx=None))
        if name_dict.keys().__contains__(""):
            del name_dict[""]
        return name_dict

    # 返回员工的响应总时长
    def get_total_response(self, name):
        print("正在查询" + name + "的响应总时长")
        m = self.col_index('处理人')
        p = self.col_index('完成时间')
        q = self.col_index('派单时间')
        name_list = list(self.table.col_values(m, start_rowx=1, end_rowx=None))
        finish_list = []
        send_list = []
        for i in range(1, self.nrow):
            cell1 = self.table.cell_value(i, p)
            if cell1 != "":
                finish_list.append(str(datetime(*xldate_as_tuple(cell1, 0)).strftime('%Y/%m/%d %H:%M:%S')))
            else:
                finish_list.append(" ")
            cell2 = self.table.cell_value(i, q)
            if cell2 != "":
                send_list.append(str(datetime(*xldate_as_tuple(cell2, 0)).strftime('%Y/%m/%d %H:%M:%S')))
            else:
                send_list.append(" ")
        print(finish_list)
        print(send_list)
        total_sec = 0
        for i in range(len(name_list)):
            if name_list[i] == name:
                if finish_list[i] != " " and send_list[i] != " ":
                    total_sec += self.minus_time_in_str(send_list[i], finish_list[i])
        return total_sec

    # 返回员工的按时解决事件数
    def get_num_solved_ontime(self, name):
        print("正在查询" + name + "的按时解决事件总数")
        m = self.col_index('处理人')
        p = self.col_index('销单时间')
        q = self.col_index('完成时间')
        name_list = list(self.table.col_values(m, start_rowx=1, end_rowx=None))
        print(name_list)
        cancel_list = []
        finish_list = []
        for i in range(1, self.nrow):
            cell1 = self.table.cell_value(i, p)
            if cell1 != "":
                cancel_list.append(str(datetime(*xldate_as_tuple(cell1, 0)).strftime('%Y/%m/%d %H:%M:%S')))
            else:
                cancel_list.append(" ")
            cell2 = self.table.cell_value(i, q)
            if cell2 != "":
                finish_list.append(str(datetime(*xldate_as_tuple(cell2, 0)).strftime('%Y/%m/%d %H:%M:%S')))
            else:
                finish_list.append(" ")
        solved_limited_list = self.get_solved_limited_list()
        print(solved_limited_list)
        on_time_num = 0
        for i in range(len(name_list)):
            if name_list[i] == name:
                if cancel_list[i] != " " and finish_list[i] != " ":
                    if self.minus_time_in_str(finish_list[i], cancel_list[i]) <= 3600 * solved_limited_list[i]:
                        on_time_num += 1
                else:
                    on_time_num += 1
        return on_time_num

    # 返回员工“成功解决”的事件总数
    def get_num_all_solved(self, name):
        print("正在查询: " + name + "的成功解决事件总数")
        m = self.col_index('处理人')
        n = self.col_index('结束代码')
        name_list = list(self.table.col_values(m, start_rowx=1, end_rowx=None))
        code_list = list(self.table.col_values(n, start_rowx=1, end_rowx=None))
        # 遍历行
        solved_num = 0
        for i in range(len(name_list)):
            if name_list[i] == name and code_list[i] == '根本解决':
                solved_num += 1
        print("solved_num: " + str(solved_num))
        return solved_num

    # 返回员工“事件解决”的总时长（处理时间(小时)总和）
    def get_total_solved_time(self, name):
        print("正在查询：" + name + "的事件总解决时长")
        m = self.col_index('处理人')
        n = self.col_index('处理时间(小时)')
        name_list = list(self.table.col_values(m, start_rowx=1, end_rowx=None))
        time_list = list(self.table.col_values(n, start_rowx=1, end_rowx=None))
        total_time = 0
        for i in range(len(name_list)):
            if name_list[i] == name and time_list[i] != "":
                total_time += float(time_list[i])
        return total_time

    # 返回事件的紧急程度列表
    def get_solved_limited_list(self):
        solved_limited_list = []
        n = self.col_index('事件优先级')
        for i in range(1, self.nrow):
            value = self.table.cell_value(i, n)
            if value == "低":
                solved_limited_list.append(72)
            elif value == "中":
                solved_limited_list.append(48)
            elif value == "高":
                solved_limited_list.append(8)
            elif value == "紧急":
                solved_limited_list.append(4)
            else:
                solved_limited_list.append(72)
        return solved_limited_list

    # 计算两个字符串时间('%Y/%m/%d %H:%M')的时间差: str2 - str1
    # 返回时间间隔, 单位: s
    @staticmethod
    def minus_time_in_str(str1, str2):
        if str1 != " " and str2 != " ":
            time1 = strptime(str1, '%Y/%m/%d %H:%M:%S')
            time2 = strptime(str2, '%Y/%m/%d %H:%M:%S')
            return mktime(time2) - mktime(time1)
        else:
            return 0

    # 返回列名返回列索引
    def col_index(self, col_name):
        first_col_list = self.table.row_values(0)  # 第一行元素生成列表
        try:
            i = first_col_list.index(col_name)
        except ValueError:
            return -1
        else:
            return i


# 排序类
class SortEngine:
    def __init__(self):
        self.dic_py = dict()
        self.dic_bh = dict()
        # 建立拼音辞典
        with open('./py.txt', 'r', encoding='utf8') as f:
            content_py = f.readlines()

            for i in content_py:
                i = i.strip()
                word_py, mean_py = i.split('\t')
                self.dic_py[word_py] = mean_py

        # 建立笔画辞典
        with open('./bh.txt', 'r', encoding='utf8') as f:
            content_bh = f.readlines()

            for i in content_bh:
                i = i.strip()
                word_bh, mean_bh = i.split('\t')
                self.dic_bh[word_bh] = mean_bh

    ###############################
    # 辞典查找函数
    def searchdict(self, dic, uchar):
        if u'\u4e00' <= uchar <= u'\u9fa5':
            value = dic.get(uchar)
            if value == None:
                value = '*'
        else:
            value = uchar
        return value

    # 比较单个字符
    def comp_char_PY(self, A, B):
        if A == B:
            return -1
        pyA = self.searchdict(self.dic_py, A)
        pyB = self.searchdict(self.dic_py, B)

        # 比较拼音
        if pyA > pyB:
            return 1
        elif pyA < pyB:
            return 0

        # 比较笔画
        else:
            bhA = eval(self.searchdict(self.dic_bh, A))
            bhB = eval(self.searchdict(self.dic_bh, B))
            if bhA > bhB:
                return 1
            elif bhA < bhB:
                return 0
            else:
                return "拼音相同，笔画也相同？"

    # 比较字符串
    def comp_char(self, A, B):

        n = min(len(A), len(B))
        i = 0
        while i < n:
            dd = self.comp_char_PY(A[i], B[i])
            # 如果第一个单词相等，就继续比较下一个单词
            if dd == -1:
                i = i + 1
                # 如果比较到头了
                if i == n:
                    dd = len(A) > len(B)
            else:
                break
        return dd

    # 排序函数
    def cnsort(self, nline):
        n = len(nline)
        lines = "\n".join(nline)

        for i in range(1, n):  # 插入法
            tmp = nline[i]
            j = i
            while j > 0 and self.comp_char(nline[j - 1], tmp):
                nline[j] = nline[j - 1]
                j -= 1
            nline[j] = tmp
        return nline


MyGUI()  # 启动窗口
