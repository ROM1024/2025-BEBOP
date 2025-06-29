日程管理系统项目 README

一、项目概述
本项目是一个集成了桌面应用与 Web API 的日程管理系统，旨在帮助用户高效地管理日程安排。桌面应用提供了直观的日历界面，方便用户查看和编辑日程。同时，系统借助大语言模型（LLM）对下周日程进行自动优化，以提高日程安排的合理性和效率。本项目为北京大学2025年春程序设计实习课上的大作业。

二、项目结构
BEBOP/
├── .env                # 环境变量配置文件，存储 DeepSeek API 密钥
├── .gitignore          # Git 忽略文件，指定不需要纳入版本控制的文件和目录
├── caption.py          # 实现日历事件管理器的 GUI 界面
├── flask_app.py        # 提供 Flask Web API 服务
├── main.py             # 项目入口文件，启动 Flask API 服务和日历事件管理器 GUI
├── 记录.xlsx           # 用于存储日程数据的 Excel 文件
└── schedules.json      # 临时存储日程数据的 JSON 文件
└── requirements.txt    #运行需要的依赖

三、功能特性
日程管理：用户可以通过桌面应用方便地查看、添加和编辑日程安排。
日历展示：以日历形式直观展示每天的日程任务，便于用户快速了解日程分布。
日程优化：利用大语言模型对下周日程进行自动优化，避免时间冲突，合理分配时间。
数据存储：日程数据持久化存储在 Excel 文件中，方便管理和备份。

四、安装与运行
下载源代码
1. 克隆项目
git clone <项目仓库地址>
cd BEBOP

2. 创建并激活虚拟环境（可选但推荐）
bash
python -m venv venv
# Linux/Mac
source venv/bin/activate
# Windows
.\venv\Scripts\activate

3. 安装依赖
pip install -r requirements.txt

4. 配置环境变量
在根目录新建 .env 文件中配置 DeepSeek API 密钥：
DEEPSEEK_API_KEY=sk-
请将上述密钥替换为你自己的有效硅基流动 DeepSeek API 密钥。

5. 运行项目
python main.py

6. 访问应用
桌面应用：启动后会弹出日历事件管理器窗口，用户可直接使用。

下载预编译版本
1. 点击下方链接下载：
(https://github.com/你的用户名/你的项目名/releases/download/v1.0.0/main.exe)

2. 下载后直接运行 `BEBOP.exe`

注意：完全不能保证程序中的API还能正常运行，且此时记录.xlsx会生成在BEBOP.exe的目录下，生成文件后要先输入一行内容才会正常运行，所以只作为试用。

五、使用说明
桌面应用
查看日程：在日历中选择日期，右侧将显示该日的日程安排。
添加日程：点击相应日期，弹出添加日程对话框，输入日程信息并保存。
优化日程：点击 “自动调整下周日程” 按钮，系统将调用 LLM 对下周日程进行优化，并更新日程安排。

六、注意事项
请确保已正确配置 DEEPSEEK_API_KEY 环境变量，否则日程优化功能将无法使用。
日程数据存储在 记录.xlsx 文件中，请确保该文件具有读写权限。
项目使用了 Flask 和 Tkinter 库，请确保系统已安装相应的依赖。

七、贡献与反馈
如果你对本项目有任何建议或发现了问题，请在 GitHub 上提交 Issue 或 Pull Request。

演示文档：
https://disk.pku.edu.cn/link/AAB9EB65B5F6314C56820CC587771131DD
