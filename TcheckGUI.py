import tkinter as tk

import xlrd
from splinter import Browser
import selenium  # 需要使用包中的异常告警：selenium.common.exceptions.InvalidElementStateException


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
        self.label_input = tk.Label(self.window, text='当月指标所在列：', width=15)
        # 创建一个输入框,用来输入指标所在的列数
        self.input_col = tk.Entry(self.window, width=10)

        # 检查excel文件中的指标的button
        self.btn = tk.Button(self.window, text='检查指标', width=8, height=2, command=self.check_count)

        # 打开web、填写指标button
        self.fill_btn = tk.Button(self.window, text='填指标', width=8, height=2, command=self.fill_web_pages)

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
        self.btn.grid(column=2, row=3)
        self.fill_btn.grid(column=3, row=3)
        for child in self.window.winfo_children():
            child.grid_configure(padx=5, pady=5)

    def check_count(self):  # 检查excel文件中的指标值
        # 读入通信能力报表数据表
        wkbk = xlrd.open_workbook('汇总表.xls')
        sheet1 = wkbk.sheet_by_name("通信能力月报表")
        sheet2 = wkbk.sheet_by_name("区域情况统计表（三）")
        sheet3 = wkbk.sheet_by_name('地市填写——移动通信能力季报表（二）')
        col_str = self.input_col.get()  # 取得当月数据所在的列（数字）

        col = 5  # 默认为第5列
        try:
            col = int(col_str)
        except ValueError:
            print('输入的是：%d' % col)
            # col = 5  # 默认为第5列
            self.info.set('请输入正确的数字！')
            return

        row_count1 = sheet1.nrows  # 总行数
        row_count2 = sheet2.nrows  # 总行数
        row_count3 = sheet3.nrows  # 总行数

        value1 = {}  # "通信能力月报表"指标值
        value2 = {}  # "区域情况统计表（三）"指标值
        value3 = {}  # '地市填写——移动通信能力季报表（二）'指标值

        # 读入通信能力报表逻辑检查公式（只导入了两个sheet，没有导入季报的审核公式）
        workbook_logic = xlrd.open_workbook('逻辑审核关系表.xls')
        sheet_logic1 = workbook_logic.sheet_by_name('月报审核公式')
        sheet_logic2 = workbook_logic.sheet_by_name('区域统计表审核公式')
        sheet_logic3 = workbook_logic.sheet_by_name('季报审核公式')

        # python审核公式所在的列。注意：这里默认在第四列！！！！！
        col_python_func = 4 - 1
        row_count_l1 = sheet_logic1.nrows  # 公式的行数
        row_count_l2 = sheet_logic2.nrows  # 公式的行数
        row_count_l3 = sheet_logic3.nrows  # 公式的行数
        func_check = []  # 审核公式数组

        # 以下两个循环将公式读入数组funcCheck
        for r in range(1, row_count_l1):
            func_check.append(sheet_logic1.cell(r, col_python_func).value)  # .value表示单元格中的值

        for r in range(1, row_count_l2):
            func_check.append(sheet_logic2.cell(r, col_python_func).value)  # .value表示单元格中的值

        for r in range(1, row_count_l3):
            func_check.append(sheet_logic3.cell(r, col_python_func).value)  # .value表示单元格中的值

        try:
            # 将"通信能力月报表"指标值读入value1字典，将所有指标定义为以指标名为变量名的变量，并为其赋值
            for r in range(4, row_count1):
                if sheet1.cell(r, 0).value != "" and sheet1.cell(r, col).value != "":
                    namer = sheet1.cell(r, 0).value
                    # 如果前面输入的col（指标所在列的值超出范围，下面这句会报错：IndexError: array index out of range
                    valuer = sheet1.cell(r, col - 1).value
                    value1[namer] = valuer
                    # eval(namer+' = '+str(valuer))  # 将字符串表示的表达式转换为表达式，并求值
                    define_var_command = compile(namer + ' = ' + str(valuer), '', 'single')  # 生成变量赋值语句
                    exec(define_var_command)  # 执行变量赋值语句

            # 将"区域情况统计表（三）"指标值读入value2字典，将所有指标定义为以指标名为变量名的变量，并为其赋值
            for r in range(4, row_count2):
                if sheet2.cell(r, 0).value != "" and sheet2.cell(r, col - 1).value != "":
                    namer = sheet2.cell(r, 0).value
                    valuer = sheet2.cell(r, col - 1).value
                    value2[namer] = valuer
                    define_var_command = compile(namer + ' = ' + str(valuer), '', 'single')  # 生成变量赋值语句
                    exec(define_var_command)  # 执行变量赋值语句

            # 将"地市填写——移动通信能力季报表（二）"指标值读入value3字典，将所有指标定义为以指标名为变量名的变量，并为其赋值
            for r in range(4, row_count3):
                if sheet3.cell(r, 0).value != "" and sheet3.cell(r, col - 1).value != "":
                    namer = sheet3.cell(r, 0).value  # 指标名称
                    valuer = sheet3.cell(r, col - 1).value  # 指标值
                    value3[namer] = valuer
                    define_var_command = compile(namer + ' = ' + str(valuer), '', 'single')  # 生成变量赋值语句
                    exec(define_var_command)  # 执行变量赋值语句
        except IndexError:
            print('输入的是：%d' % col)
            self.info.set('输入的数字超出了范围！请重新输入：')
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

        return

    def fill_web_pages(self):
        # 读入通信能力报表数据表
        wkbk = xlrd.open_workbook('汇总表.xls')
        sheet1 = wkbk.sheet_by_name("通信能力月报表")
        sheet2 = wkbk.sheet_by_name("区域情况统计表（三）")
        sheet3 = wkbk.sheet_by_name('地市填写——移动通信能力季报表（二）')
        col_str = self.input_col.get()  # 取得当月数据所在的列（数字）

        col = 5  # 默认当月数据在第五列！！！！！！！！！！！！！
        try:
            col = int(col_str)
        except ValueError:
            print('输入的是：%d' % col)
            # col = 5  # 默认为第5列
            self.info.set('请输入正确的数字！')
            return

        row_count1 = sheet1.nrows  # 总行数
        row_count2 = sheet2.nrows  # 总行数
        row_count3 = sheet3.nrows  # 总行数

        value1 = {}  # "通信能力月报表"指标值
        value2 = {}  # "区域情况统计表（三）"指标值
        value3 = {}  # "地市填写——移动通信能力季报表（二）"指标值

        reason1 = {}  # "通信能力月报表"指标变更原因
        reason2 = {}  # "区域情况统计表（三）"指标变更原因
        reason3 = {}  # "地市填写——移动通信能力季报表（二）"指标变更原因

        try:
            # 将"通信能力月报表"指标值读入value1字典
            for r in range(4, row_count1):  # r:当前行
                if sheet1.cell(r, 0).value != "" and sheet1.cell(r, col - 1).value != "":
                    namer = sheet1.cell(r, 0).value
                    valuer = int(sheet1.cell(r, col - 1).value)
                    reason = sheet1.cell(r, col).value  # 变更原因（值在col-1列（列数从0开始计数），原因在col列）
                    value1[namer] = valuer  # 指标值
                    if reason == '':
                        reason1[namer] = '  '
                    else:
                        reason1[namer] = reason

            # 将"区域情况统计表（三）"指标值读入value2字典
            for r in range(4, row_count2):  # r:当前行
                if sheet2.cell(r, 0).value != "" and sheet2.cell(r, col - 1).value != "":
                    namer = sheet2.cell(r, 0).value
                    valuer = int(sheet2.cell(r, col - 1).value)
                    reason = sheet2.cell(r, col).value  # 变更原因
                    value2[namer] = valuer
                    if reason == '':
                        reason2[namer] = '  '
                    else:
                        reason2[namer] = reason

            # 将"地市填写——移动通信能力季报表（二）"指标值读入value3字典
            for r in range(4, row_count3):  # r:当前行
                if sheet3.cell(r, 0).value != "" and sheet3.cell(r, col - 1).value != "":
                    namer = sheet3.cell(r, 0).value
                    valuer = int(sheet3.cell(r, col - 1).value)
                    reason = sheet3.cell(r, col).value  # 变更原因
                    value3[namer] = valuer
                    if reason == '':
                        reason3[namer] = '  '
                    else:
                        reason3[namer] = reason
        except IndexError:
            print('输入的是：%d' % col)
            self.info.set('输入的数字超出了范围！请重新输入：')
            return

        # 整合成一个value字典
        # value = {}
        # value.update(value1)
        # value.update(value2)

        # 使用chrome浏览器
        # 如果是windows且chrome未安装在默认位置，则需指定安装位置
        # executable_path = {'executable_path': 'C:\Program Files\Google\Chrome\Application\chrome.exe'}
        browser = Browser('chrome')  # , **executable_path)
        # 访问通信能力月报表网站
        browser.visit('http://10.204.51.174/txnl/Login.aspx')
        # 填写用户名、密码
        browser.find_by_xpath('//*[@id="txtUserName"]')
        browser.fill('txtUserName', 'yh')
        browser.fill('txtPassword', 'yh')
        # 点击登陆按钮
        browser.find_by_xpath('//*[@id="ibtnSubmit"]').first.click()

        # ----------------------------填报"移动通信能力月报表"---------------------------
        # 在左侧iframe中操作
        with browser.get_iframe('ltbfrm') as ltbfrm:
            # 点击页面左侧的"移动通信能力月报表"，在右半部分的iframe中打开填报页面
            browser.find_by_xpath('//*[@id="ProjectManagement0101"]/table[1]/tbody/tr/td[2]').first.click()

        # 在右侧iframe中操作
        browser.is_element_present_by_id('rtmfrm', wait_time=10)
        with browser.get_iframe('rtmfrm') as rtmfrm:
            # 填"移动通信能力月报表"时的所在行计数器(共179行）
            for row in range(4, 180):  # 右侧指标表格最多180行
                rtmfrm.is_element_present_by_id('td' + str(row) + '_0', wait_time=10)  # 判断对应的网页元素是否已经显示出来
                key = browser.find_by_id('td' + str(row) + '_0').first.value  # 获取网页当前行等指标名称
                # key=browser.find_by_xpath('//*[@id="td4_0"]').first.value
                if key in value1:
                    print('当前指标：{0}。'.format(key))
                    browser.find_by_id('td' + str(row) + '_3').click()  # 点击指标对应的"太原"列的单元格，弹出指标填写对话框
                    # str = browser.find_by_xpath('//*[@id="txt_zbmc"]').value  # 读取"指标名称"
                    try:  # 如果网页元素不存在，则说明暂时还不能填写（此时selenium会报错：selenium.common.exceptions.InvalidElementStateException）
                        browser.find_by_id('txt_reason').fill(reason1[key])  # 填入指标修改原因
                        browser.find_by_id('txt_xgz').fill(value1[key])  # 填入指标值
                        browser.find_by_id('btn_tj').click()  # 点击提交按钮
                    except selenium.common.exceptions.InvalidElementStateException:
                        self.info.set('指标暂时不能填写！')
                        return

                    alert = browser.get_alert()  # 点击提交按钮后网也会弹出alert对话框要求确认
                    alert.accept()
                    # print(alert.text)

        # ---------------------------填报"移动通信能力月报表"---------------------------

        # ---------------------------填报"区域情况统计表（三）"---------------------------
        with browser.get_iframe('ltbfrm') as ltbfrm:
            # 点击页面左侧的"区域情况统计表（三）"，在右半部分的iframe中打开填报页面
            browser.find_by_xpath('//*[@id="ProjectManagement0101"]/table[5]/tbody/tr/td[2]').first.click()

        # 在右侧iframe中操作
        browser.is_element_present_by_id('rtmfrm', wait_time=20)
        with browser.get_iframe('rtmfrm') as rtmfrm:
            # 填"区域情况统计表（三）"时的所在行计数器(共104行）
            for row in range(4, 105):  # 右侧指标表格最多105行
                rtmfrm.is_element_present_by_id('td' + str(row) + '_0', wait_time=10)
                key = browser.find_by_id('td' + str(row) + '_0').value  # 获取指标名称
                if key in value2:
                    browser.find_by_id('td' + str(row) + '_3').click()  # 点击指标对应的"太原"列的单元格，弹出指标填写对话框
                    # str = browser.find_by_xpath('//*[@id="txt_zbmc"]').value  # 读取"指标名称"
                    try:
                        browser.find_by_id('txt_reason').fill(reason2[key])
                        browser.find_by_id('txt_xgz').fill(value2[key])
                        browser.find_by_id('btn_tj').click()
                    except selenium.common.exceptions.InvalidElementStateException:
                        self.info.set('指标暂时不能填写！')
                        return

                    alert = browser.get_alert()  # 点击提交按钮后网也会弹出alert对话框要求确认
                    alert.accept()
                    # print('当前指标：{0}。'.format(key))

        # ---------------------------填报"区域情况统计表（三）"---------------------------

        if len(value3) == 0:  # 如果季度报表中的指标值均为空，则不填季报，直接退出
            return
            # ---------------------------填报"地市填写——移动通信能力季报表（二）"---------------------------
        with browser.get_iframe('ltbfrm') as ltbfrm:
            # 点击页面左侧的"地市填写——移动通信能力季报表（二）"，在右半部分的iframe中打开填报页面
            browser.find_by_xpath('//*[@id="ProjectManagement0101"]/table[21]/tbody/tr/td[2]').first.click()

        # 在右侧iframe中操作
        browser.is_element_present_by_id('rtmfrm', wait_time=20)
        with browser.get_iframe('rtmfrm') as rtmfrm:
            # 填"区域情况统计表（三）"时的所在行计数器(共104行）
            for row in range(4, 191):  # 右侧指标表格最多191行
                rtmfrm.is_element_present_by_id('td' + str(row) + '_0', wait_time=10)
                key = browser.find_by_id('td' + str(row) + '_0').value  # 获取指标名称
                if key in value2:
                    browser.find_by_id('td' + str(row) + '_3').click()  # 点击指标对应的"太原"列的单元格，弹出指标填写对话框
                    try:
                        browser.find_by_id('txt_reason').fill(reason2[key])
                        browser.find_by_id('txt_xgz').fill(value2[key])
                        browser.find_by_id('btn_tj').click()
                    except selenium.common.exceptions.InvalidElementStateException:
                        self.info.set('指标暂时不能填写！')
                        return
                    alert = browser.get_alert()  # 点击提交按钮后网也会弹出alert对话框要求确认
                    alert.accept()

                    # ---------------------------填报"地市填写——移动通信能力季报表（二）"---------------------------


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
