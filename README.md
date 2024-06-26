工具简介

本项目实现了一个基于PyQt5的图形化自动化模拟点击的工具，通过配置xpath等页面查找元素的语句来实现自动化登陆并操作任意系统的操作，完全支持自定义的各类点击及用户输入流程，且可完美匹配iframe，遇到iframe也可切换到iframe中进行元素匹配，理论上可以实现全部的网页自动化操作，极大提升了通用性，简化了Web自动化操作流程。以下是主要功能介绍：

一、用户界面

🔗 URL输入框：用户可以输入需要访问的网页地址。

📋 操作表格：使用QTreeView展示三列表格，分别是操作模式（Mode）、路径（Path）和数据（Data）。

➕ 添加与删除：用户可以通过“Add”按钮添加新行，通过上下文菜单删除选中的行。

📁 导入与导出：支持通过按钮将配置导入或导出为Excel文件。


二、自动化操作

🌐 Selenium WebDriver：使用Selenium WebDriver实现自动化浏览器操作。

🔍 多种定位方式：支持多种元素定位方式（如xpath、id、name等），根据用户输入的路径和数据执行操作。

🔒 验证码处理：下载验证码图片，使用ddddocr库识别并自动填入验证码。


三、多线程

⏳ 后台执行：使用QThread实现Selenium操作的后台执行，避免阻塞主界面。

🎛️ 控制执行：提供开始和停止按钮，控制Selenium自动化脚本的执行。

四、错误处理和通知

❗ 异常捕捉：在执行过程中捕捉各种异常，并通过消息框提示用户。


五、导入导出配置


📊 Excel文件：可以将操作步骤和URL导出为Excel文件，并支持从Excel文件导入配置。


✨ 关键功能点

🖥️ 图形用户界面：使用PyQt5创建，包含输入框、按钮、树视图等控件。

🌐 自动化浏览器操作：基于Selenium WebDriver，支持多种元素定位方式。

🔒 验证码识别：使用ddddocr库识别网页中的验证码图片。

📁 配置管理：支持导入和导出操作配置文件（Excel格式）。

⏳ 多线程处理：使用QThread避免界面阻塞，提供友好的用户体验。

这个工具帮助用户通过图形界面轻松配置和运行复杂的Web自动化任务，同时支持处理网页验证码，极大地方便了自动化操作的实现和管理。

工具使用

环境准备

1、第一步，安装python3、pip3(安装python3的时候自带)

2、安装pip3依赖（运行如下命令安装）

```markdown
pip3 install -r requirements.txt
```

3、Google或Microsoft官网下载对应你当前浏览器版本的driver驱动，下载好后文件需要解压到AutoClick.py文件目录下，脚本默认从当前目录加载驱动。

```markdown
# 下载好后在如下代码位置进行修改，修改为你下载的对应的driver名字即可
service = Service(executable_path="msedgedriver.exe")
```

4、启动应用程序

运行如下命令启动应用程序

```markdown
python3 AutoClick.py
```

5、界面展示


![image](https://github.com/BoldChampion/AutoClick/assets/171965684/3b31e2bd-f856-4321-add0-73223ce6b063)


![image](https://github.com/BoldChampion/AutoClick/assets/171965684/168b03ff-f6df-4843-ad7e-c88b4439a45e)


希望这个工具能帮到大家，让大家提升工作效率，避免重复繁杂的工作浪费大家太多的时间，如有任何针对工具的想法，大家可以评论区多多提意见，我会在后续不断优化工具，提升工具的全面性。

完全开源，代码随便拿去用，大家可自行魔改，如果帮到大家的话请一键三连，给我点点小星星✨✨✨✨✨✨✨✨✨

‼️‼️‼️ 重要申明 ‼️‼️‼️

此脚本仅限代码学习使用，切勿用于任何违法乱纪类项目，如有违反与本人无关。

