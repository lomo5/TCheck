# TCheck
## 功能：
- 检查"汇总表.xls"中的指标是否符合"逻辑审核关系表.xls"中的逻辑关系公式。
- 打开浏览器并登陆网站，将"汇总表.xls"中的指标值自动填入。
## 主程序TcheckGUI.py实现：
- 基于xlrd读取excel文档
- 基于Splinter操纵Chrome浏览器
- 使用TKinter实现图形界面
- 使用了selenium的错误消息
## 类TcheckGUI(object)的方法：
- gui_arrange(self)： 进行所有控件的布局。
- check_count(self)： 检查excel文件中的指标值。
- __get_values_assignment_commands(self, sheet, column)： 从sheet中读取指标值，并生成指标变量赋值语句存入一个list中。
- __get_counts_and_reasons(sheet, column)： 获取指标sheet中的指标值和修改原因，返回对象：指标值dict、原因dict。（@staticmethod）
- fill_web_pages(self, sheet_num)： 打开浏览器，调用__fill_web()填入指标。
- __fill_web(self, browser, xpath_left, rows, values, reasons)： 填入指标。
## 使用方法：
1. 安装chrome浏览器。
2. 将chromedriver.exe（附件中有）的目录加入系统环境变量PATH参数中。
3. 将指标填到"汇总表.xls"中。
    - 注意：不能修改这两个文件的名称、内部表结构。
    - 当月指标默认为"汇总表.xls"中每个sheet的第5列，最好不要修改，如果要修改必须同时修改所有sheet。
4. 将指标excel文件："汇总表.xls"和"逻辑审核关系表.xls"与本程序置于同一目录下。
5. 打开程序点击“检查指标”按钮，可以检查"汇总表.xls"中填入的指标是否符合逻辑审核关系中的逻辑关系。
6. 如果检查没有问题，点击后面的按钮，将自动打开Chrome浏览器，填写对应的报表。
7. 为确保填写无误，程序不会自动提交审核，填报完成后需要在浏览器页面手工提交审核。

## 笔记:
1. 如果页面加载较慢需要等待：browser.is_element_present_by_id('rtmfrm', wait_time=10)，有例子中使用time.sleep(8) 来实现等待（未测试过）
2. 如果引用的页面元素不存在则回报错：selenium.common.exceptions.InvalidElementStateException。注意使用该异常，需要import selenium。
3. windows中，如果chrome安装位置不是默认位置则需指定安装位置：
```
executable_path = {'executable_path': 'C:\Program Files\Google\Chrome\Application\chrome.exe'}
browser = Browser('chrome', **executable_path)
```
4. 如何安装[chrome driver](http://splinter.readthedocs.io/en/latest/drivers/chrome.html)。windows下载了driver之后需要将chromedriver.exe所在目录加入系统环境变量PATH中。
5. html中的<td>内的<span>标签在读取<td>内容时（browser.find_by_id('td' + str(row) + '_3').first.value）直接被忽略
6. splinter用来检测页面元素是否已经加载的is_element_present_by_id等函数的可选参数wait_time的单位是秒。如果检测到元素已加载则立即返回True（即使时间未到），否则等待相应的时间。