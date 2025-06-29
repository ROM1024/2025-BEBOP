import threading
import tkinter as tk
from caption import CalendarApp
from flask_app import app
import os
from dotenv import load_dotenv

# 加载 .env 文件
load_dotenv()

def run_flask_app():
    """运行 Flask API 服务"""
    print("启动 Flask API 服务...")
    app.run(host='0.0.0.0', port=5000, debug=True, use_reloader=False)

def run_calendar_app():
    """运行日历事件管理器 GUI"""
    print("启动日历事件管理器...")
    root = tk.Tk()
    app = CalendarApp(root)  # 实例化 CalendarApp 类
    root.mainloop()

if __name__ == "__main__":
    # 从环境变量中获取 API 密钥
    api_key = os.getenv('DEEPSEEK_API_KEY')
    if not api_key:
        print("警告: 未找到 DEEPSEEK_API_KEY 环境变量，优化功能可能受限")
    
    # 创建并启动 Flask 线程
    flask_thread = threading.Thread(target=run_flask_app, daemon=True)
    flask_thread.start()
    
    # 在主线程中运行 Tkinter 应用
    run_calendar_app()
    
    print("程序已退出")