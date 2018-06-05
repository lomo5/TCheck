import xlrd

# import time

# thismonth = time.strftime('%m', time.localtime(time.time()))
# thisyear = time.strftime('%Y', time.localtime(time.time()))
# print('%s年%s月' % (thisyear, thismonth))

# 读入通信能力报表数据表
wkbk = xlrd.open_workbook('check.xls')
sheet1 = wkbk.sheet_by_name("通信能力月报表")
sheet2 = wkbk.sheet_by_name("区域情况统计表（三）")
colStr = input("请输入当月数据所在的列（数字）：")
if colStr == '':
    col = 5  # 默认为第5列
else:
    col = int(colStr)
print(col)
row_count1 = sheet1.nrows  # 总行数
row_count2 = sheet2.nrows  # 总行数
value1 = {}  # "通信能力月报表"指标值
value2 = {}  # "区域情况统计表（三）"指标值

# 读入通信能力报表逻辑检查公式（只导入了两个sheet，没有导入季报的审核公式）
wkbk_logic = xlrd.open_workbook('逻辑审核关系表.xls')
sheet_logic1 = wkbk_logic.sheet_by_name('月报审核公式')
sheet_logic2 = wkbk_logic.sheet_by_name('区域统计表审核公式')

# python审核公式所在的列。注意：这里默认在第四列！！！！！
colPythonFunc = 4 - 1
rowCountL1 = sheet_logic1.nrows  # 行数
rowCountL2 = sheet_logic2.nrows  # 行数
funcCheck = []  # 审核公式数组

# 以下两个循环将公式读入数组funcCheck
for r in range(1, rowCountL1):
    funcCheck.append(sheet_logic1.cell(r, colPythonFunc).value)  # .value表示单元格中的值

for r in range(1, rowCountL2):
    funcCheck.append(sheet_logic2.cell(r, colPythonFunc).value)  # .value表示单元格中的值

# 将"通信能力月报表"指标值读入value1字典
for r in range(4, row_count1):
    if sheet1.cell(r, 0).value != "":
        namer = sheet1.cell(r, 0).value
        valuer = sheet1.cell(r, col - 1).value
        value1[namer] = valuer
        # eval(namer+' = '+str(valuer))
        fuzhi = compile(namer + ' = ' + str(valuer), '', 'single')  # 生成变量赋值语句
        exec(fuzhi)  # 执行变量赋值语句

# 将"区域情况统计表（三）"指标值读入value2字典
for r in range(4, row_count2):
    if sheet2.cell(r, 0).value != "":
        namer = sheet2.cell(r, 0).value
        valuer = sheet2.cell(r, col - 1).value
        value2[namer] = valuer
        fuzhi = compile(namer + ' = ' + str(valuer), '', 'single')  # 生成变量赋值语句
        exec(fuzhi)  # 执行变量赋值语句

result = True  # 记录是否所有公式均为true
count = 0  # 存在的公式数
countNone = 0  # 有未定义变量的公式数
for func in funcCheck:
    # 通过NameError（变量不存在）错误来跳过不存在的变量
    try:
        resultTemp = eval(func)  # 当前公式是否为True
        if not resultTemp:  # 打印出错的公式
            print(func + ':' + str(resultTemp))
        result = result and resultTemp
        count += 1
    except NameError:
        countNone += 1
        pass  # 如果出现变量未定义报错则忽略
        # print('变量不存在：'+func)

if result:
    print('已通过所有逻辑审核公式检验，共验证%d个公式。不存在的公式%d个。' % (count, countNone))
else:
    print('未通过逻辑审核公式检验')
