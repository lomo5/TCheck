import xlrd
from splinter import Browser

# 读入通信能力报表数据表
wkbk = xlrd.open_workbook('check.xls')
sheet1 = wkbk.sheet_by_name("通信能力月报表")
sheet2 = wkbk.sheet_by_name("区域情况统计表（三）")

col = 5  # 当月数据在第五列！！！！！！！！！！！！！
row_count1 = sheet1.nrows  # 总行数
row_count2 = sheet2.nrows  # 总行数
value1 = {}  # "通信能力月报表"指标值
value2 = {}  # "区域情况统计表（三）"指标值
reason1 = {}  # "通信能力月报表"指标变更原因
reason2 = {}  # "区域情况统计表（三）"指标变更原因

# 将"通信能力月报表"指标值读入value1字典
for r in range(4, row_count1):
    if sheet1.cell(r, 0).value != "":
        namer = sheet1.cell(r, 0).value
        valuer = int(sheet1.cell(r, col - 1).value)
        reason = sheet1.cell(r, col).value  # 变更原因
        value1[namer] = valuer  # 指标值
        if reason == '':
            reason1[namer] = '  '
        else:
            reason1[namer] = reason

# 将"区域情况统计表（三）"指标值读入value2字典
for r in range(4, row_count2):
    if sheet2.cell(r, 0).value != "":
        namer = sheet2.cell(r, 0).value
        valuer = int(sheet2.cell(r, col - 1).value)
        reason = sheet2.cell(r, col).value  # 变更原因
        value2[namer] = valuer
        if reason == '':
            reason2[namer] = '  '
        else:
            reason2[namer] = reason

# 整合成一个value字典
value = {}
value.update(value1)
value.update(value2)

# 使用chrome浏览器
browser = Browser('chrome')
# 访问通信能力月报表网站
browser.visit('http://10.204.51.174/txnl/Login.aspx')
# 填写用户名、密码
browser.find_by_xpath('//*[@id="txtUserName"]')
browser.fill('txtUserName', 'yh')
browser.fill('txtPassword', 'yh')
# 点击登陆按钮
browser.find_by_xpath('//*[@id="ibtnSubmit"]').first.click()

# vvvvvvvvvvvvvvvvvvvvvvvvvvv填报"移动通信能力月报表"vvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
# 在左侧iframe中操作
with browser.get_iframe('ltbfrm') as ltbfrm:
    # 点击页面左侧的"移动通信能力月报表"，在右半部分的iframe中打开填报页面
    browser.find_by_xpath('//*[@id="ProjectManagement0101"]/table[1]/tbody/tr/td[2]').first.click()

# 在右侧iframe中操作
browser.is_element_present_by_id('rtmfrm', wait_time=10)
with browser.get_iframe('rtmfrm') as rtmfrm:
    # 填"移动通信能力月报表"时的所在行计数器(共179行）
    for row in range(4, 180):
        rtmfrm.is_element_present_by_id('td' + str(row) + '_0', wait_time=10)
        key = browser.find_by_id('td' + str(row) + '_0').first.value # 获取指标名称
        # key=browser.find_by_xpath('//*[@id="td4_0"]').first.value
        if key in value1:
            print('当前指标：{0}。'.format(key))
            browser.find_by_id('td' + str(row) + '_3').click()  # 点击指标对应的"太原"列的单元格，弹出指标填写对话框
            # str = browser.find_by_xpath('//*[@id="txt_zbmc"]').value  # 读取"指标名称"
            browser.find_by_id('txt_reason').fill(reason1[key])
            browser.find_by_id('txt_xgz').fill(value1[key])
            browser.find_by_id('btn_tj').click()
            alert=browser.get_alert()
            alert.accept()
            print(alert.text)


# ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^填报"移动通信能力月报表"^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


# vvvvvvvvvvvvvvvvvvvvvvvvvvvvvv填报"区域情况统计表（三）"vvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
with browser.get_iframe('ltbfrm') as ltbfrm:
    # 点击页面左侧的"区域情况统计表（三）"，在右半部分的iframe中打开填报页面
    browser.find_by_xpath('//*[@id="ProjectManagement0101"]/table[5]/tbody/tr/td[2]').first.click()

# 在右侧iframe中操作
browser.is_element_present_by_id('rtmfrm', wait_time=20)
with browser.get_iframe('rtmfrm') as rtmfrm:
    # 填"区域情况统计表（三）"时的所在行计数器(共104行）
    for row in range(4, 105):
        rtmfrm.is_element_present_by_id('td' + str(row) + '_0', wait_time=10)
        key = browser.find_by_id('td' + str(row) + '_0').value  # 获取指标名称
        if key in value2:
            browser.find_by_id('td' + str(row) + '_3').click()  # 点击指标对应的"太原"列的单元格，弹出指标填写对话框
            # str = browser.find_by_xpath('//*[@id="txt_zbmc"]').value  # 读取"指标名称"
            browser.find_by_id('txt_reason').fill(reason2[key])
            browser.find_by_id('txt_xgz').fill(value2[key])
            browser.find_by_id('btn_tj').click()
            alert = browser.get_alert()
            alert.accept()
        # print('当前指标：{0}。'.format(key))

# ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^填报"区域情况统计表（三）"^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
