# TCheck
## check.py
- 检查check.xls中的指标是否符合“逻辑审核关系表”中的逻辑关系公式。
- 给予xlrd读取excel文档
## fillwebpages.py
- 打开浏览器并登陆网站，将check.xls中的指标值自动填入。
- 基于Splinter操纵chrome浏览器
## TcheckGUI.py
- 图形界面版本，包含check.py和fillwebpages.py的功能；
- 基于tkinter实现图形界面
- 使用了selenium的错误消息

## 笔记:
1. 如果页面加载较慢需要等待：browser.is_element_present_by_id('rtmfrm', wait_time=10)，有例子中使用time.sleep(8) 来实现等待（未测试过）
2. 如果引用的页面元素不存在则回报错：selenium.common.exceptions.InvalidElementStateException。注意使用该异常，需要import selenium。
3. windows中，如果chrome安装位置不是默认位置则需指定安装位置：
'''
executable_path = {'executable_path': 'C:\Program Files\Google\Chrome\Application\chrome.exe'}
browser = Browser('chrome', **executable_path)
'''
4. 如何安装[chrome driver](http://splinter.readthedocs.io/en/latest/drivers/chrome.html)。windows下载了driver之后需要将chromedriver.exe所在目录加入系统环境变量PATH中。
