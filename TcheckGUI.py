import tkinter as tk

import selenium  # 需要使用包中的异常告警：selenium.common.exceptions.InvalidElementStateException
import xlrd
from splinter import Browser


class TcheckGUI(object):
    # 初始化类
    def __init__(self):
        # 创建主窗口,用于容纳其它组件
        self.window = tk.Tk()
        self.window.title('通信能力报表自动填报程序')
        # 窗口尺寸，注意宽、高中间是小写字母x！！！
        self.window.geometry()
        # self.mainframe = tk.Frame(self.window,  relief='flat', bg='green')

        # 以下创建各控件
        # 设置label内容及格式，其中width和height的单位是单个字符的宽和高
        self.label1 = tk.Label(self.window,
                               text='1、本机必须安装chrome浏览器。\n'
                                    + '2、请将chromedriver.exe的目录加入系统环境变量PATH参数中。\n'
                                    + '3、请将指标excel文件："汇总表.xls"和"逻辑审核关系表.xls"与本程序置于同一目录下。且不能修改这两个文档的名称、结构等内容。\n'
                                    + '4、如果本月不提交季报，请务必将季报指标值所在列（默认为第五列）清空！\n'
                                    + '5、填报完成后需要手工提交审核。'
                                    + '6、本月指标默认为每个sheet的第五列。', justify='left')  # ,width=50, height=2, bg='green')
        # 信息显示的label
        self.info = tk.StringVar()
        self.info.set('请填入check.xls中，本月指标所在的列（数字）。')  # 此变量绑定到label_info
        self.label_info = tk.Label(self.window, textvariable=self.info, bg='limegreen')  # , bg='blue')
        # 输入框前的提示：
        self.label_input = tk.Label(self.window, text='当月指标所在列（默认第5列）：', width=30)
        # 创建一个输入框,用来输入指标所在的列数
        self.input_col = tk.Entry(self.window, width=10)

        # 检查excel文件中的指标的button
        self.btn = tk.Button(self.window, text='检查指标', width=10, height=2, command=self.check_count)

        # 打开web、填写指标button。移动通信能力月报表：1，"区域情况统计表（三）"：2，"地市填写——移动通信能力季报表（二）"
        self.fill_btn1 = tk.Button(self.window, text='填通信能力月报表', width=20, height=2,
                                   command=lambda: self.fill_web_pages(sheet_num=1))
        self.fill_btn2 = tk.Button(self.window, text='填区域情况统计表', width=20, height=2,
                                   command=lambda: self.fill_web_pages(sheet_num=2))
        self.fill_btn3 = tk.Button(self.window, text='填地市通信能力季报表', width=25, height=2,
                                   command=lambda: self.fill_web_pages(sheet_num=3))

    def gui_arrange(self):  # 进行所有控件的布局
        # self.mainframe.grid(column=0, row=0, sticky="N W E S")
        #  为什么frame不能设置底色？A：是可以的，不能如网上的例子调用ttk，而是要调用tk来创建
        #  详见：https://blog.csdn.net/nkd50000/article/details/77511707?locationNum=2&fps=1
        # self.mainframe.columnconfigure(0, weight=1)
        # self.mainframe.rowconfigure(0, weight=1)
        self.label1.grid(column=0, row=0, columnspan=4, sticky='W')
        self.label_info.grid(column=0, row=1, columnspan=4, sticky='W', )
        self.label_input.grid(column=0, row=2, sticky='E')
        self.input_col.grid(column=1, row=2, sticky='W')
        self.btn.grid(column=0, row=3)
        self.fill_btn1.grid(column=1, row=3)
        self.fill_btn2.grid(column=2, row=3)
        self.fill_btn3.grid(column=3, row=3)
        for child in self.window.winfo_children():
            child.grid_configure(padx=5, pady=5)

    def check_count(self):  # 检查excel文件中的指标值
        # 读入通信能力报表数据表
        try:
            wkbk = xlrd.open_workbook('汇总表.xls')
        except FileNotFoundError:
            self.info.set('"汇总表.xls"文件不存在！')
            return
        sheet1 = wkbk.sheet_by_name("通信能力月报表")
        sheet2 = wkbk.sheet_by_name("区域情况统计表（三）")
        sheet3 = wkbk.sheet_by_name('地市填写——移动通信能力季报表（二）')
        col_str = self.input_col.get()  # 取得当月数据所在的列（数字）

        try:
            col = int(col_str)
        except ValueError:
            col = 5  # 如果输入非数字，或为空，则默认为第5列

        # 读入通信能力报表逻辑检查公式
        try:
            workbook_logic = xlrd.open_workbook('逻辑审核关系表.xls')
        except FileNotFoundError:
            self.info.set('"逻辑审核关系表.xls"文件不存在！')
            return
        sheet_logic1 = workbook_logic.sheet_by_name('月报审核公式')
        sheet_logic2 = workbook_logic.sheet_by_name('区域统计表审核公式')
        sheet_logic3 = workbook_logic.sheet_by_name('季报审核公式')

        # python审核公式所在的列。注意：这里默认在第四列！！！！！
        col_python_func = 4 - 1
        row_count_l1 = sheet_logic1.nrows  # 公式的行数
        row_count_l2 = sheet_logic2.nrows  # 公式的行数
        row_count_l3 = sheet_logic3.nrows  # 公式的行数
        func_check = []  # 审核公式数组

        # 以下3个循环将公式读入数组funcCheck
        for r in range(1, row_count_l1):
            func_check.append(sheet_logic1.cell(r, col_python_func).value)  # .value表示单元格中的值

        for r in range(1, row_count_l2):
            func_check.append(sheet_logic2.cell(r, col_python_func).value)  # .value表示单元格中的值

        for r in range(1, row_count_l3):
            func_check.append(sheet_logic3.cell(r, col_python_func).value)  # .value表示单元格中的值

        commands = []  # 指标变量赋值指令列表
        try:
            # 读取三个sheet中的指标值，生成赋值命令list
            commands += self.__get_values_assignment_commands(sheet1, col)  # "通信能力月报表"
            commands += self.__get_values_assignment_commands(sheet2, col)  # "区域情况统计表（三）"
            commands += self.__get_values_assignment_commands(sheet3, col)  # "地市填写——移动通信能力季报表（二）"
            for comm in commands:
                exec(comm)  # 执行变量赋值语句
        except IndexError:
            print('输入的是：%d' % col)
            self.info.set('输入的数字超出了范围！请重新输入：')
            return
        except SyntaxError:
            self.info.set('输入的列值不正确！请重新输入：')
            return

        result = True  # 记录是否所有公式均为true
        count = 0  # 存在的公式数
        count_none = 0  # 有未定义变量的公式数
        str_wrong = ''  # 出错的指标
        for func in func_check:
            # 通过NameError（变量不存在）错误来跳过不存在的变量
            try:
                result_temp = eval(func)  # 当前公式是否为True
                if not result_temp:  # 打印出错的公式
                    print(func + ':' + str(result_temp))
                    str_wrong += '\n' + func + ';'
                result = result and result_temp
                count += 1
            except NameError:
                count_none += 1
                pass  # 如果出现变量未定义报错则忽略
                # print('变量不存在：'+func)

        if result:
            print('已通过所有逻辑审核公式检验，共验证%d个公式。不存在的公式%d个。' % (count, count_none))
            self.info.set('已通过所有逻辑审核公式检验，共验证%d个公式。不存在的公式%d个。' % (count, count_none))
        else:
            print('未通过逻辑审核公式检验')
            self.info.set('未通过逻辑审核公式:%s' % str_wrong)
        # return

    # 从sheet中读取指标值，并生成指标变量赋值语句存入一个list中
    def __get_values_assignment_commands(self, sheet, column):
        # sheet:excel表格的sheet对象
        # column：指标值所在列
        # 返回：赋值语句对象的列表

        row_count = sheet.nrows  # 总行数
        values = {}  # 指标值字典
        command_list = []
        for r in range(4, row_count):
            # 如果指标值名称（第一列）为空，或者该指标的值未填（默认第5列），则忽略该指标。
            if sheet.cell(r, 0).value != "" and sheet.cell(r, column - 1).value != "":
                namer = sheet.cell(r, 0).value  # 指标名称
                # 如果前面输入的col（指标所在列的值超出范围，下面这句会报错：IndexError: array index out of range
                valuer = sheet.cell(r, column - 1).value  # 指标值
                values[namer] = valuer
                # eval(namer+' = '+str(valuer))  # 将字符串表示的表达式转换为表达式，并求值
                define_var_command = compile(namer + ' = ' + str(valuer), '', 'single')  # 生成变量赋值语句对象
                command_list.append(define_var_command)
        return command_list

    # 获取指标sheet中的指标值和修改原因，返回对象：指标值dict、原因dict
    @staticmethod
    def __get_counts_and_reasons(sheet, column):
        # sheet:excel表格的sheet对象
        # column：指标值所在列

        row_count = sheet.nrows  # 总行数
        values = {}  # 指标值字典
        reasons = {}  # 原因字典
        for r in range(4, row_count):  # r:当前行
            if sheet.cell(r, 0).value != "" and sheet.cell(r, column - 1).value != "":
                namer = sheet.cell(r, 0).value
                valuer = int(sheet.cell(r, column - 1).value)
                reason = sheet.cell(r, column).value  # 变更原因（值在col-1列（列数从0开始计数），原因在col列）
                values[namer] = valuer  # 指标值
                if reason == '':
                    reasons[namer] = '  '
                else:
                    reasons[namer] = reason
        return values, reasons

    # 填表
    def fill_web_pages(self, sheet_num):
        # 读入通信能力报表数据表
        # sheet_num:表示统计表对应的值，移动通信能力月报表：1，"区域情况统计表（三）"：2，"地市填写——移动通信能力季报表（二）"：3
        try:
            wkbk = xlrd.open_workbook('汇总表.xls')
        except FileNotFoundError:
            self.info.set('"汇总表.xls"文件不存在！')
            return
        sheet1 = wkbk.sheet_by_name("通信能力月报表")
        sheet2 = wkbk.sheet_by_name("区域情况统计表（三）")
        sheet3 = wkbk.sheet_by_name('地市填写——移动通信能力季报表（二）')
        col_str = self.input_col.get()  # 取得当月数据所在的列（数字）

        try:
            col = int(col_str)
        except ValueError:
            col = 5  # 如果没有输入值，或者输入非数字，则默认为第5列

        try:
            value1, reason1 = self.__get_counts_and_reasons(sheet1, col)  # 将"通信能力月报表"指标值、原因读入字典
            value2, reason2 = self.__get_counts_and_reasons(sheet2, col)  # 将"区域情况统计表（三）"指标值、原因读入字典
            value3, reason3 = self.__get_counts_and_reasons(sheet3, col)  # 将"地市填写——移动通信能力季报表（二）"指标值、原因读入字典
        except IndexError:
            print('输入的是：%d' % col)
            self.info.set('输入的数字超出了范围！请重新输入：')
            return
        except ValueError:
            self.info.set('输入的列值不正确！请重新输入：')
            return

        ''' 
        使用chrome浏览器
        如果是windows且chrome未安装在默认位置，则需指定安装位置
        executable_path = {'executable_path': 'C:\Program Files\Google\Chrome\Application\chrome.exe'}
        '''
        browser = Browser('chrome')  # , **executable_path)
        # 访问通信能力月报表网站
        browser.visit('http://10.204.51.174/txnl/Login.aspx')
        # 填写用户名、密码
        browser.find_by_xpath('//*[@id="txtUserName"]')
        browser.fill('txtUserName', 'yh')
        browser.fill('txtPassword', 'yh')
        # 点击登陆按钮
        browser.find_by_xpath('//*[@id="ibtnSubmit"]').first.click()

        # 填入指标、原因：
        try:  # 如果网页元素不存在，则说明暂时还不能填写（此时selenium会报错：selenium.common.exceptions.InvalidElementStateException）
            if sheet_num == 1:  # 是1，表示点击了第一个按钮，要填"移动通信能力月报表"
                # 填报"移动通信能力月报表"
                self.__fill_web(browser, '//*[@id="ProjectManagement0101"]/table[1]/tbody/tr/td[2]', 180, value1,
                                reason1)
            elif sheet_num == 2:
                # 填报"区域情况统计表（三）"
                self.__fill_web(browser, '//*[@id="ProjectManagement0101"]/table[5]/tbody/tr/td[2]', 105, value2,
                                reason2)
            else:
                if len(value3) == 0:  # 如果季度报表中的指标值均为空，则不填季报，直接退出
                    return
                # 填报"地市填写——移动通信能力季报表（二）"
                self.__fill_web(browser, '//*[@id="ProjectManagement0101"]/table[21]/tbody/tr/td[2]', 191, value3,
                                reason3)
        except selenium.common.exceptions.InvalidElementStateException:
            self.info.set('指标暂时不能填写！')
            return

        # ----------------------------填报"移动通信能力月报表"---------------------------
        # 在左侧iframe中操作
        # with browser.get_iframe('ltbfrm') as ltbfrm:
        #     # 点击页面左侧的"移动通信能力月报表"，在右半部分的iframe中打开填报页面
        #     browser.find_by_xpath('//*[@id="ProjectManagement0101"]/table[1]/tbody/tr/td[2]').first.click()
        #
        # # 在右侧iframe中操作
        # browser.is_element_present_by_id('rtmfrm', wait_time=10)
        # with browser.get_iframe('rtmfrm') as rtmfrm:
        #     # 填"移动通信能力月报表"时的所在行计数器(共179行）
        #     for row in range(4, 180):  # 右侧指标表格最多180行
        #         rtmfrm.is_element_present_by_id('td' + str(row) + '_0', wait_time=10)  # 判断对应的网页元素是否已经显示出来
        #         key = browser.find_by_id('td' + str(row) + '_0').first.value  # 获取网页当前行等指标名称
        #         # key=browser.find_by_xpath('//*[@id="td4_0"]').first.value
        #         if key in value1:
        #             print('当前指标：{0}。'.format(key))
        #             browser.find_by_id('td' + str(row) + '_3').click()  # 点击指标对应的"太原"列的单元格，弹出指标填写对话框
        #             # str = browser.find_by_xpath('//*[@id="txt_zbmc"]').value  # 读取"指标名称"
        #             try:  # 如果网页元素不存在，则说明暂时还不能填写（此时selenium会报错：selenium.common.exceptions.InvalidElementStateException）
        #                 browser.find_by_id('txt_reason').fill(reason1[key])  # 填入指标修改原因
        #                 browser.find_by_id('txt_xgz').fill(value1[key])  # 填入指标值
        #                 browser.find_by_id('btn_tj').click()  # 点击提交按钮
        #             except selenium.common.exceptions.InvalidElementStateException:
        #                 self.info.set('指标暂时不能填写！')
        #                 return
        #
        #             alert = browser.get_alert()  # 点击提交按钮后网也会弹出alert对话框要求确认
        #             alert.accept()
        #             # print(alert.text)
        # # ---------------------------填报"移动通信能力月报表"---------------------------
        #
        # # ---------------------------填报"区域情况统计表（三）"---------------------------
        # with browser.get_iframe('ltbfrm') as ltbfrm:
        #     # 点击页面左侧的"区域情况统计表（三）"，在右半部分的iframe中打开填报页面
        #     browser.find_by_xpath('//*[@id="ProjectManagement0101"]/table[5]/tbody/tr/td[2]').first.click()
        #
        # # 在右侧iframe中操作
        # browser.is_element_present_by_id('rtmfrm', wait_time=20)
        # with browser.get_iframe('rtmfrm') as rtmfrm:
        #     # 填"区域情况统计表（三）"时的所在行计数器(共104行）
        #     for row in range(4, 105):  # 右侧指标表格最多105行
        #         rtmfrm.is_element_present_by_id('td' + str(row) + '_0', wait_time=10)
        #         key = browser.find_by_id('td' + str(row) + '_0').value  # 获取指标名称
        #         if key in value2:
        #             browser.find_by_id('td' + str(row) + '_3').click()  # 点击指标对应的"太原"列的单元格，弹出指标填写对话框
        #             # str = browser.find_by_xpath('//*[@id="txt_zbmc"]').value  # 读取"指标名称"
        #             try:
        #                 browser.find_by_id('txt_reason').fill(reason2[key])
        #                 browser.find_by_id('txt_xgz').fill(value2[key])
        #                 browser.find_by_id('btn_tj').click()
        #             except selenium.common.exceptions.InvalidElementStateException:
        #                 self.info.set('指标暂时不能填写！')
        #                 return
        #
        #             alert = browser.get_alert()  # 点击提交按钮后网也会弹出alert对话框要求确认
        #             alert.accept()
        #             # print('当前指标：{0}。'.format(key))
        #
        # # ---------------------------填报"区域情况统计表（三）"---------------------------
        #
        # if len(value3) == 0:  # 如果季度报表中的指标值均为空，则不填季报，直接退出
        #     return
        #     # ---------------------------填报"地市填写——移动通信能力季报表（二）"---------------------------
        # with browser.get_iframe('ltbfrm') as ltbfrm:
        #     # 点击页面左侧的"地市填写——移动通信能力季报表（二）"，在右半部分的iframe中打开填报页面
        #     browser.find_by_xpath('//*[@id="ProjectManagement0101"]/table[21]/tbody/tr/td[2]').first.click()
        #
        # # 在右侧iframe中操作
        # browser.is_element_present_by_id('rtmfrm', wait_time=20)
        # with browser.get_iframe('rtmfrm') as rtmfrm:
        #     # 填"区域情况统计表（三）"时的所在行计数器(共104行）
        #     for row in range(4, 191):  # 右侧指标表格最多191行
        #         rtmfrm.is_element_present_by_id('td' + str(row) + '_0', wait_time=10)
        #         key = browser.find_by_id('td' + str(row) + '_0').value  # 获取指标名称
        #         if key in value2:
        #             browser.find_by_id('td' + str(row) + '_3').click()  # 点击指标对应的"太原"列的单元格，弹出指标填写对话框
        #             try:
        #                 browser.find_by_id('txt_reason').fill(reason2[key])
        #                 browser.find_by_id('txt_xgz').fill(value2[key])
        #                 browser.find_by_id('btn_tj').click()
        #             except selenium.common.exceptions.InvalidElementStateException:
        #                 self.info.set('指标暂时不能填写！')
        #                 return
        #             alert = browser.get_alert()  # 点击提交按钮后网也会弹出alert对话框要求确认
        #             alert.accept()

        # ---------------------------填报"地市填写——移动通信能力季报表（二）"---------------------------

    def __fill_web(self, browser, xpath_left, rows, values, reasons):
        # browser：浏览器对象
        # xpath_left:左侧菜单中相应链接的xpath
        # rows：页面右侧表格最大行数
        # values：指标值dict
        # reasons：原因dict

        # 在左侧iframe中操作
        with browser.get_iframe('ltbfrm') as ltbfrm:
            # 点击页面左侧的"移动通信能力月报表"，在右半部分的iframe中打开填报页面
            browser.find_by_xpath(xpath_left).first.click()

        # 在右侧iframe中操作
        browser.is_element_present_by_id('rtmfrm', wait_time=10)
        with browser.get_iframe('rtmfrm') as rtmfrm:
            # row：所在行计数器(共rows行）
            for row in range(4, rows):  # 右侧指标表格最多rows行
                rtmfrm.is_element_present_by_id('td' + str(row) + '_0', wait_time=10)  # 判断对应的网页元素是否已经显示出来
                key = browser.find_by_id('td' + str(row) + '_0').first.value  # 获取网页当前行的指标名称
                if key in values:  # 指标值dict中有该值，没有则忽略
                    filled_key = browser.find_by_id('td' + str(row) + '_3').first.value  # 获取当前该指标的值
                    if filled_key == '':  # 该指标还未填时才填
                        print('当前指标：{0}。'.format(key))
                        browser.find_by_id('td' + str(row) + '_3').click()  # 点击指标对应的"太原"列的单元格，弹出指标填写对话框
                        # 此处不拦截selenium.common.exceptions.InvalidElementStateException异常，留到上层函数中处理！
                        browser.find_by_id('txt_reason').fill(reasons[key])  # 填入指标修改原因
                        browser.find_by_id('txt_xgz').fill(values[key])  # 填入指标值
                        browser.find_by_id('btn_tj').click()  # 点击提交按钮
                        alert = browser.get_alert()  # 点击提交按钮后网也会弹出alert对话框要求确认
                        alert.accept()


def main():
    # 初始化对象
    t_chk = TcheckGUI()
    # 进行布局
    t_chk.gui_arrange()

    # 执行主程序
    # t_chk.window.mainloop() # 这个也可以
    tk.mainloop()
    pass


if __name__ == '__main__':
    main()
