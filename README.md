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
