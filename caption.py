import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from calendar import monthcalendar, month_name, day_name
from datetime import datetime, timedelta
import pandas as pd
import os
import re
import tkinter.simpledialog as sd
import json  # 添加JSON支持
import sys
from flask_app import LLMAPI  # 从flask_app.py中导入LLMAPI类

if getattr(sys, 'frozen', False):
    # 打包后的环境：使用 sys.executable 获取 exe 路径
    BASE_DIR = os.path.dirname(sys.executable)  # main.exe 所在目录
else:
    # 开发环境：使用 __file__ 获取脚本路径
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# 构建 Excel 文件路径（保存到 exe 同级目录）
EXCEL_FILE_PATH = os.path.join(BASE_DIR, "记录.xlsx")

class CalendarApp:
    def __init__(self, root):
        self.root = root
        self.root.title("日历事件管理器")
        self.root.geometry("1100x800")  # 增加宽度以适应新列
        self.root.resizable(True, True)

        self.api_key = os.getenv('DEEPSEEK_API_KEY')
        self.llm_api = LLMAPI(self.api_key) if self.api_key else None  # 使用flask_app.py中的LLMAPI类
        self.SCHEDULE_FILE = "schedules.json"
        self.FEEDBACK_FILE = "feedbacks.json"

        # 存储事件的数据结构 {日期: [{"time": "", "task": "", "completion": ""}, ...]}
        self.events = {}

        # 当前显示的日期
        self.current_date = datetime.now()

        # 尝试从Excel加载事件
        self.load_events_from_excel()

        # 创建UI
        self.create_widgets()
        self.update_calendar()

        # 设置关闭窗口事件处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.modified = False  # 跟踪是否有未保存的修改

    def adjust_next_week_schedule(self):
        try:
            # 1. 从Excel加载所有事件
            self.load_events_from_excel()

            # 2. 获取下周日期范围
            today = datetime.now().date()
            next_monday = today + timedelta(days=(7 - today.weekday()))
            next_sunday = next_monday + timedelta(days=6)

            # 3. 提取下周所有事件
            next_week_events = {}
            current_date = next_monday
            while current_date <= next_sunday:
                date_str = current_date.strftime("%Y-%m-%d")
                next_week_events[date_str] = self.events.get(date_str, [])
                current_date += timedelta(days=1)

            total_events = sum(len(events) for events in next_week_events.values())
            if total_events == 0:
                messagebox.showinfo("提示", "下周没有安排任何事件，无需调整")
                return

            print(f"找到下周 {len(next_week_events)} 天的 {total_events} 个事件")

            # 4. 调用LLM进行优化
            optimized_events = self.optimize_with_llm(next_week_events)

            if not optimized_events:
                messagebox.showerror("失败", "无法自动调整下周日程")
                return

            # 5. 更新内存中的事件并保存到Excel
            for date_str, events in optimized_events.items():
                self.events[date_str] = events

            if self.save_events_to_excel():
                # 更新日历显示
                self.update_calendar()
                # 如果当前正在查看下周，刷新显示
                if self.selected_date in optimized_events:
                    self.show_events(self.context_row, self.context_col)

                messagebox.showinfo("成功", "下周日程已自动调整并保存")
            else:
                messagebox.showerror("失败", "保存调整后的日程失败")
        except Exception as e:
            print(f"调整下周日程时出错: {str(e)}")
            messagebox.showerror("错误", f"调整下周日程时出错: {str(e)}")

    def optimize_with_llm(self, events_data):
        """使用LLM优化事件安排"""
        try:
            if not self.llm_api:
                messagebox.showerror("错误", "未配置API密钥，无法使用优化功能")
                return None

            # 准备LLM提示（移除了多余的缩进）
            prompt = f"""你是一个专业的日程优化顾问，请根据以下事件安排优化下周日程：

## 原始日程安排
{json.dumps(events_data, indent=2, ensure_ascii=False)}

## 优化要求
1. 保持每天的核心事件不变
2. 优化时间分配，避免冲突
3. 确保重要任务有足够时间
4. 添加必要的休息时间
5. 保持总事件数量大致相同
6. 时间格式统一为"HH:MM-HH:MM"
7. 对于已完成的事件，保持原样不变
8. 对于未开始的事件，可以调整时间
9. 如果一天任务太多，可以将任务调到之后的日期
## 输出要求
返回优化后的完整日程JSON对象，格式必须严格如下:
{{
  "2024-06-10": [
    {{"time": "09:00-10:00", "task": "会议", "completion": "待评价"}},
    {{"time": "10:30-12:00", "task": "项目开发", "completion": "待评价"}}
  ],
  "2024-06-11": [
    // 其他日期...
  ]
}}
只返回纯JSON，不要包含任何解释性文字或额外内容。"""

            # 调用LLM
            print("请求LLM优化日程...")
            response = self.llm_api.generate_response(prompt, max_tokens=2000)
            print(f"LLM响应: {response[:500]}...")

            # 更健壮的JSON解析方法
            optimized_events = self.extract_json_from_response(response)
            if not optimized_events:
                # 记录详细的响应内容以便调试
                print(f"无效的LLM响应: {response}")
                return None

            # 验证格式
            if not self.validate_optimized_events(optimized_events):
                print(f"验证失败的优化结果: {optimized_events}")
                return None

            return optimized_events

        except ValueError as ve:
            # 处理JSON解析错误
            print(f"JSON解析错误: {str(ve)}")
            return None
        except AttributeError as ae:
            # 处理API调用错误
            print(f"API调用错误: {str(ae)}")
            return None
        except Exception as e:
            # 处理其他未知错误
            print(f"LLM优化失败: {str(e)}")
            return None

    def extract_json_from_response(self, response):
        """从LLM响应中提取JSON"""
        try:
            # 尝试找到JSON开始和结束位置
            start_idx = response.find('{')
            end_idx = response.rfind('}')

            if start_idx == -1 or end_idx == -1:
                print("未找到有效的JSON结构")
                return None

            json_str = response[start_idx:end_idx+1]
            return json.loads(json_str)
        except Exception as e:
            print(f"解析JSON时出错: {str(e)}")
            return None

    
    def extract_json_from_response(self, response):
        """从LLM响应中提取JSON"""
        try:
            # 尝试找到JSON开始和结束位置
            start_idx = response.find('{')
            end_idx = response.rfind('}')
            
            if start_idx == -1 or end_idx == -1:
                print("未找到有效的JSON结构")
                return None
            
            json_str = response[start_idx:end_idx+1]
            return json.loads(json_str)
        except Exception as e:
            print(f"解析JSON失败: {str(e)}")
            return None
    
    def validate_optimized_events(self, events):
        """验证优化后的事件格式"""
        if not isinstance(events, dict):
            print("优化后事件格式错误: 应为字典")
            return False
        
        for date_str, event_list in events.items():
            # 验证日期格式
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                print(f"无效日期格式: {date_str}")
                return False
            
            # 验证事件列表
            if not isinstance(event_list, list):
                print(f"{date_str} 的事件格式错误: 应为列表")
                return False
                
            for event in event_list:
                if not all(key in event for key in ["time", "task", "completion"]):
                    print(f"事件缺少必要字段: {event}")
                    return False
                
                # 验证时间格式
                if not re.match(r"\d{2}:\d{2}-\d{2}:\d{2}", event["time"]):
                    print(f"时间格式错误: {event['time']} (应为HH:MM-HH:MM)")
                    return False
                    
        return True

    def create_widgets(self):
        # 创建主框架
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左侧日历区域
        left_frame = tk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 右侧事件区域
        right_frame = tk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        # ========== 日历区域 ==========
        # 顶部控制栏
        control_frame = tk.Frame(left_frame, padx=10, pady=10)
        control_frame.pack(fill=tk.X)
        
        self.month_var = tk.StringVar()
        self.year_var = tk.StringVar()
        
        # 月份选择
        tk.Label(control_frame, text="月份:").pack(side=tk.LEFT)
        month_combo = ttk.Combobox(control_frame, textvariable=self.month_var, width=10)
        month_combo['values'] = [month_name[i] for i in range(1, 13)]
        month_combo.current(self.current_date.month - 1)
        month_combo.pack(side=tk.LEFT, padx=5)
        month_combo.bind("<<ComboboxSelected>>", self.update_calendar)
        
        # 年份选择
        tk.Label(control_frame, text="年份:").pack(side=tk.LEFT, padx=(10, 0))
        year_combo = ttk.Combobox(control_frame, textvariable=self.year_var, width=6)
        year_combo['values'] = [str(year) for year in range(2020, 2031)]
        year_combo.set(str(self.current_date.year))
        year_combo.pack(side=tk.LEFT, padx=5)
        year_combo.bind("<<ComboboxSelected>>", self.update_calendar)
        
        # 操作按钮
        button_frame = tk.Frame(control_frame)
        button_frame.pack(side=tk.RIGHT)
        
        tk.Button(button_frame, text="今天", command=self.show_today).pack(side=tk.LEFT, padx=2)
        tk.Button(button_frame, text="保存", command=self.save_events).pack(side=tk.LEFT, padx=2)
        tk.Button(button_frame, text="加载", command=self.load_events).pack(side=tk.LEFT, padx=2)
        # 添加新按钮
        tk.Button(button_frame, text="自动调整下周日程", command=self.adjust_next_week_schedule).pack(side=tk.LEFT, padx=2)
        
        # 日历显示区域
        calendar_frame = tk.Frame(left_frame, relief=tk.GROOVE, borderwidth=2)
        calendar_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 星期标题
        weekdays = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
        for i, day in enumerate(weekdays):
            tk.Label(calendar_frame, text=day, font=("Arial", 10, "bold"), 
                     relief=tk.RAISED, padx=10, pady=5).grid(row=0, column=i, sticky="nsew")
        
        # 创建日历网格 (6行 x 7列)
        self.day_buttons = []
        for row in range(1, 7):
            row_buttons = []
            for col in range(7):
                btn = tk.Button(calendar_frame, text="", height=2, width=5,
                                command=lambda r=row, c=col: self.show_events(r, c))
                btn.grid(row=row, column=col, sticky="nsew", padx=2, pady=2)
                row_buttons.append(btn)
            self.day_buttons.append(row_buttons)
        
        # 设置网格权重
        for i in range(7):
            calendar_frame.columnconfigure(i, weight=1)
        for i in range(1, 7):
            calendar_frame.rowconfigure(i, weight=1)
        
        # ========== 事件区域 ==========
        # 事件详情框架
        detail_frame = tk.LabelFrame(right_frame, text="事件详情", padx=10, pady=10)
        detail_frame.pack(fill=tk.BOTH, expand=True)
        
        # 日期标题
        self.date_label = tk.Label(detail_frame, text="选择日期查看事件", font=("Arial", 12, "bold"))
        self.date_label.pack(pady=5)
        
        # 表格框架
        table_frame = tk.Frame(detail_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 表头 - 增加格式刷列
        headers = ["时间/时间段", "任务", "完成度", "格式刷"]
        for col, header in enumerate(headers):
            tk.Label(table_frame, text=header, font=("Arial", 10, "bold"), 
                     relief=tk.RAISED, padx=5, pady=5).grid(row=0, column=col, sticky="nsew")
        
        # 创建10行输入表格
        self.time_entries = []
        self.task_entries = []
        self.completion_entries = []
        self.format_brush_buttons = []  # 存储格式刷按钮
        
        for row in range(1, 11):
            # 时间输入框
            time_entry = tk.Entry(table_frame, width=12)
            time_entry.grid(row=row, column=0, padx=2, pady=2, sticky="nsew")
            self.time_entries.append(time_entry)
            time_entry.bind("<FocusOut>", self.on_event_modified)  # 添加焦点离开事件
            
            # 任务输入框
            task_entry = tk.Entry(table_frame, width=30)
            task_entry.grid(row=row, column=1, padx=2, pady=2, sticky="nsew")
            self.task_entries.append(task_entry)
            task_entry.bind("<FocusOut>", self.on_event_modified)  # 添加焦点离开事件
            
            # 完成度输入框
            completion_var = tk.StringVar()
            completion_combo = ttk.Combobox(table_frame, textvariable=completion_var, width=8)
            completion_combo['values'] = ["未开始", "进行中", "已完成", "延期", "取消"]
            completion_combo.grid(row=row, column=2, padx=2, pady=2, sticky="nsew")
            self.completion_entries.append(completion_combo)
            completion_combo.bind("<<ComboboxSelected>>", self.on_event_modified)  # 添加选择事件
            
            # 格式刷按钮
            brush_btn = tk.Button(table_frame, text="📋", width=3, 
                                 command=lambda r=row-1: self.show_format_brush_menu(r))
            brush_btn.grid(row=row, column=3, padx=2, pady=2, sticky="nsew")
            self.format_brush_buttons.append(brush_btn)
        
        # 设置表格网格权重
        for i in range(4):  # 现在有4列
            table_frame.columnconfigure(i, weight=1)
        for i in range(1, 11):
            table_frame.rowconfigure(i, weight=1)
        
        # 按钮区域 - 移除了"添加事件"按钮
        btn_frame = tk.Frame(detail_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        tk.Button(btn_frame, text="保存更改", command=self.save_current_events).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="删除事件", command=self.delete_event).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="清空事件", command=self.clear_events).pack(side=tk.LEFT, padx=5)
        
        # 当前选择的日期
        self.selected_date = None
        self.context_row = None
        self.context_col = None
        
        # 当前日历数据
        self.current_cal = None

    def parse_excel_date(self, date_str):
        """解析Excel中的日期格式(年.月.日)"""
        try:
            # 尝试解析"年.月.日"格式
            if isinstance(date_str, str) and re.match(r"\d{4}\.\d{1,2}\.\d{1,2}", date_str):
                parts = date_str.split('.')
                if len(parts) == 3:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])
                    return f"{year}-{month:02d}-{day:02d}"
            
            # 处理pandas日期类型
            if isinstance(date_str, pd.Timestamp):
                return date_str.strftime("%Y-%m-%d")
            
            # 处理datetime对象
            if isinstance(date_str, datetime):
                return date_str.strftime("%Y-%m-%d")
            
            # 处理字符串格式的日期
            if isinstance(date_str, str):
                # 尝试解析其他常见格式
                for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S"):
                    try:
                        dt = datetime.strptime(date_str, fmt)
                        return dt.strftime("%Y-%m-%d")
                    except ValueError:
                        continue
            
            # 尝试解析Excel序列号日期
            try:
                if isinstance(date_str, (int, float)):
                    # Excel日期是从1900-01-01开始的天数
                    base_date = datetime(1900, 1, 1)
                    parsed_date = base_date + timedelta(days=date_str - 2)  # Excel有1900年闰年错误
                    return parsed_date.strftime("%Y-%m-%d")
            except:
                pass
            
            return None
        except Exception as e:
            print(f"解析日期错误: {date_str} - {str(e)}")
            return None

    def format_excel_date(self, date_str):
        """将日期格式化为年.月.日格式"""
        try:
            # 解析标准日期格式
            if isinstance(date_str, str) and re.match(r"\d{4}-\d{2}-\d{2}", date_str):
                year, month, day = date_str.split('-')
                return f"{int(year)}.{int(month)}.{int(day)}"
            return date_str
        except Exception as e:
            print(f"格式化日期错误: {date_str} - {str(e)}")
            return date_str

    def normalize_time(self, time_str):
        """标准化时间格式"""
        if not isinstance(time_str, str):
            return ""
        
        # 尝试统一时间分隔符
        time_str = time_str.replace("：", ":")  # 替换中文冒号
        time_str = time_str.replace("—", "-")   # 替换中文破折号
        time_str = time_str.replace("~", "-")   # 替换波浪号
        
        # 确保时间段分隔符统一
        if "-" in time_str:
            parts = time_str.split("-")
            if len(parts) == 2:
                start = self.normalize_single_time(parts[0].strip())
                end = self.normalize_single_time(parts[1].strip())
                return f"{start} - {end}"
        
        # 处理单个时间点
        return self.normalize_single_time(time_str.strip())
    
    def normalize_single_time(self, time_str):
        """标准化单个时间点格式"""
        # 尝试解析时间格式
        if re.match(r"\d{1,2}:\d{2}", time_str):
            # 已经是标准格式
            return time_str
        
        # 尝试添加分钟部分
        if re.match(r"\d{1,2}$", time_str):
            return f"{time_str}:00"
        
        # 其他格式直接返回
        return time_str

    def time_to_minutes(self, time_str):
        """将时间字符串转换为分钟数用于排序"""
        # 处理时间段（取开始时间）
        if " - " in time_str:
            time_str = time_str.split(" - ")[0].strip()
        
        # 尝试解析时间
        try:
            if ":" in time_str:
                parts = time_str.split(":")
                hours = int(parts[0])
                minutes = int(parts[1]) if len(parts) > 1 else 0
                return hours * 60 + minutes
            elif time_str.isdigit():
                return int(time_str) * 60
            else:
                return 0
        except:
            return 0

    def load_events_from_excel(self):
        """从Excel文件加载事件"""
        try:
            if os.path.exists(EXCEL_FILE_PATH):
                # 读取Excel文件
                df = pd.read_excel(EXCEL_FILE_PATH, dtype=str)  # 全部读取为字符串
                print(f"成功读取Excel文件，共{len(df)}行数据")
                
                # 检查必要的列
                required_columns = ["日期", "时间", "任务", "完成度"]
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    print(f"Excel文件缺少必要列: {', '.join(missing_columns)}")
                    messagebox.showwarning("警告", f"Excel文件缺少必要列: {', '.join(missing_columns)}")
                    return False
                
                # 清空当前事件
                self.events = {}
                
                # 处理每一行数据
                for index, row in df.iterrows():
                    date_str = str(row["日期"])
                    parsed_date = self.parse_excel_date(date_str)
                    
                    if not parsed_date:
                        print(f"跳过无法解析的日期: {date_str} (行 {index+2})")
                        continue
                    
                    # 处理时间
                    time_str = self.normalize_time(str(row["时间"]))
                    
                    # 创建事件
                    event = {
                        "time": time_str,
                        "task": str(row["任务"]) if pd.notna(row["任务"]) else "",
                        "completion": str(row["完成度"]) if pd.notna(row["完成度"]) else "未开始"
                    }
                    
                    if parsed_date not in self.events:
                        self.events[parsed_date] = []
                    
                    self.events[parsed_date].append(event)
                    print(f"添加事件: {parsed_date} - {time_str} - {event['task']}")
                
                # 对所有日期的事件按时间排序
                for date in self.events:
                    self.events[date] = sorted(
                        self.events[date], 
                        key=lambda x: self.time_to_minutes(x["time"])
                    )
                    print(f"排序后 {date} 有 {len(self.events[date])} 个事件")
                
                print(f"成功从Excel加载 {len(df)} 条事件记录")
                return True
            else:
                print("未找到记录.xlsx文件，将使用空事件集")
                return False
        except Exception as e:
            print(f"加载Excel文件时出错: {str(e)}")
            messagebox.showerror("错误", f"加载Excel文件时出错: {str(e)}")
            return False

    def save_events_to_excel(self):
        """将事件保存到Excel文件"""
        try:
            # 准备数据
            data = []
            event_count = 0
            
            for date, events in self.events.items():
                for event in events:
                    formatted_date = self.format_excel_date(date)
                    data.append({
                        "日期": formatted_date,
                        "时间": event["time"],
                        "任务": event["task"],
                        "完成度": event["completion"]
                    })
                    event_count += 1
                    print(f"保存事件: {formatted_date} - {event['time']} - {event['task']}")
            
            # 创建数据框
            df = pd.DataFrame(data)
            
            # 确保列顺序正确
            if not df.empty:
                df = df[["日期", "时间", "任务", "完成度"]]
            
            # 保存到Excel
            df.to_excel(EXCEL_FILE_PATH, index=False)
            print(f"成功保存 {event_count} 条事件到Excel")
            return True
        except Exception as e:
            print(f"保存Excel文件时出错: {str(e)}")
            messagebox.showerror("错误", f"保存Excel文件时出错: {str(e)}")
            return False

    def update_calendar(self, event=None):
        try:
            year = int(self.year_var.get())
            month = list(month_name).index(self.month_var.get())
            
            # 获取当前月的日历
            cal = monthcalendar(year, month)
            self.current_cal = cal  # 保存当前日历数据
            
            # 重置所有按钮
            for row in self.day_buttons:
                for btn in row:
                    btn.config(text="", bg="SystemButtonFace", state=tk.NORMAL)
            
            # 填充日历
            for week_idx, week in enumerate(cal):
                for day_idx, day in enumerate(week):
                    if day != 0:
                        btn = self.day_buttons[week_idx][day_idx]
                        btn.config(text=str(day))
                        
                        # 标记有事件的日期
                        date_str = f"{year}-{month:02d}-{day:02d}"
                        if date_str in self.events and self.events[date_str]:
                            btn.config(bg="#ADD8E6")
                        # 标记今天
                        today = datetime.now()
                        if year == today.year and month == today.month and day == today.day:
                            btn.config(bg="#FFD700")
            print(f"日历更新为: {year}年{month}月")
        except Exception as e:
            print(f"更新日历时出错: {str(e)}")

    def show_today(self):
        today = datetime.now()
        self.month_var.set(month_name[today.month])
        self.year_var.set(str(today.year))
        self.update_calendar()

        year = int(self.year_var.get())
        month = list(month_name).index(self.month_var.get())
        cal = monthcalendar(year, month)
        day = today.day
        for row_idx, week in enumerate(cal):
            if day in week:
                col = week.index(day)
                row = row_idx + 1  # 注意: 行索引从1开始
                break
        self.show_events(row, col)

    def show_events(self, row, col):
        try:
            # 获取点击的日期 - 从日历数据中获取而不是按钮文本
            if not self.current_cal:
                return
                
            day = self.current_cal[row-1][col]  # 注意: row从1开始，日历数据从0开始
            if day == 0:  # 0表示非当月日期
                return
                
            month = list(month_name).index(self.month_var.get())
            year = int(self.year_var.get())
            date_str = f"{year}-{month:02d}-{day:02d}"
            self.selected_date = date_str
            
            # 保存点击位置（用于刷新）
            self.context_row = row
            self.context_col = col
            
            # 更新日期标题
            display_date = f"{year}年{month}月{day}日"
            self.date_label.config(text=f"{display_date} 事件")
            
            # 清空表格
            for time_entry in self.time_entries:
                time_entry.delete(0, tk.END)
            for task_entry in self.task_entries:
                task_entry.delete(0, tk.END)
            for completion_combo in self.completion_entries:
                completion_combo.set("")
            
            # 填充表格（按时间排序）
            if date_str in self.events:
                events = self.events[date_str]
                
                # 确保事件按时间排序
                sorted_events = sorted(events, key=lambda x: self.time_to_minutes(x["time"]))
                
                for i, event in enumerate(sorted_events):
                    if i < 10:  # 最多显示10个事件
                        self.time_entries[i].insert(0, event["time"])
                        self.task_entries[i].insert(0, event["task"])
                        self.completion_entries[i].set(event["completion"])
            print(f"显示事件: {date_str}")
        except Exception as e:
            print(f"显示事件时出错: {str(e)}")

    def save_current_events(self):
        """保存当前日期的所有事件"""
        if not self.selected_date:
            messagebox.showwarning("警告", "请先选择一个日期")
            return
        
        # 创建新的事件列表
        new_events = []
        
        # 遍历所有行
        for i in range(10):
            time_val = self.time_entries[i].get().strip()
            task_val = self.task_entries[i].get().strip()
            completion_val = self.completion_entries[i].get().strip()
            
            # 只保存非空任务
            if task_val:
                # 标准化时间
                time_val = self.normalize_time(time_val) or "全天"
                completion_val = completion_val or "未开始"
                
                new_events.append({
                    "time": time_val,
                    "task": task_val,
                    "completion": completion_val
                })
        
        # 按时间排序
        new_events = sorted(new_events, key=lambda x: self.time_to_minutes(x["time"]))
        
        # 更新内存中的事件
        if new_events:
            self.events[self.selected_date] = new_events
        elif self.selected_date in self.events:
            del self.events[self.selected_date]
        
        # 更新日历显示
        self.update_calendar()
        self.modified = True

    def delete_event(self):
        """删除选定行的事件"""
        try:
            if not self.selected_date:
                messagebox.showwarning("警告", "请先选择一个日期")
                return
            
            # 获取选定行
            selected_row = None
            for i, entry in enumerate(self.task_entries):
                if entry.get().strip():
                    selected_row = i
                    break
            
            if selected_row is None:
                messagebox.showwarning("警告", "没有可删除的事件")
                return
            
            # 删除事件
            if self.selected_date in self.events and selected_row < len(self.events[self.selected_date]):
                del self.events[self.selected_date][selected_row]
                
                # 如果没有事件了，删除日期键
                if not self.events[self.selected_date]:
                    del self.events[self.selected_date]
                
                # 更新UI
                self.show_events(self.context_row, self.context_col)
                self.update_calendar()
                self.modified = True
            else:
                messagebox.showwarning("警告", "找不到要删除的事件")
        except Exception as e:
            print(f"删除事件时出错: {str(e)}")

    def clear_events(self):
        try:
            if not self.selected_date:
                messagebox.showwarning("警告", "请先选择一个日期")
                return
            
            if self.selected_date in self.events:
                del self.events[self.selected_date]
            
            # 更新UI
            self.show_events(self.context_row, self.context_col)
            self.update_calendar()
            self.modified = True
        except Exception as e:
            print(f"清空事件时出错: {str(e)}")

    def save_events(self):
        if self.save_events_to_excel():
            self.modified = False
            messagebox.showinfo("成功", "事件已保存到记录.xlsx")

    def load_events(self):
        if self.load_events_from_excel():
            self.update_calendar()
            if self.selected_date:
                self.show_events(self.context_row, self.context_col)
            messagebox.showinfo("成功", "事件已从记录.xlsx加载")

    def on_closing(self):
        """窗口关闭时的事件处理"""
        if self.modified:
            if messagebox.askyesno("未保存的更改", "有未保存的更改，是否保存到Excel?"):
                self.save_events_to_excel()
        self.root.destroy()
        
    def on_event_modified(self, event=None):
        """当事件被修改时调用"""
        self.modified = True
        # 自动保存当前日期的更改（不显示提示）
        self.save_current_events()
        
    def show_format_brush_menu(self, row_index):
        """显示格式刷菜单"""
        if not self.selected_date:
            messagebox.showwarning("警告", "请先选择一个日期")
            return
            
        # 检查该行是否有事件
        if not self.task_entries[row_index].get().strip():
            messagebox.showwarning("警告", "该行没有事件")
            return
            
        # 创建菜单
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="按星期", command=lambda: self.apply_format_brush(row_index, "weekly"))
        menu.add_command(label="单双周", command=lambda: self.apply_format_brush(row_index, "biweekly"))
        menu.add_command(label="按日期", command=lambda: self.apply_format_brush(row_index, "daily"))
        
        # 显示菜单
        menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())
        
    def apply_format_brush(self, row_index, mode):
        """应用格式刷"""
        try:
            # 获取当前事件信息
            time_val = self.time_entries[row_index].get().strip()
            task_val = self.task_entries[row_index].get().strip()
            
            if not task_val:
                messagebox.showwarning("警告", "该行没有事件")
                return
                
            # 获取当前日期
            current_date = datetime.strptime(self.selected_date, "%Y-%m-%d")
            
            # 根据模式获取目标日期
            target_dates = []
            
            if mode == "weekly":
                # 按星期：复制到指定的星期几
                weekdays = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
                selected = sd.askstring("按星期复制", "选择星期几(用逗号分隔, 如: 1,3,5)\n1:周一 2:周二 ... 7:周日", 
                                      initialvalue=str(current_date.isoweekday()))
                
                if not selected:
                    return
                    
                try:
                    days = [int(d.strip()) for d in selected.split(",") if d.strip()]
                    for day in days:
                        if day < 1 or day > 7:
                            raise ValueError("星期值必须在1-7之间")
                except Exception as e:
                    messagebox.showerror("错误", f"无效的输入: {str(e)}")
                    return
                
                # 计算未来四周内指定的星期几
                for week_offset in range(1, 5):  # 未来4周
                    for day in days:
                        # 计算目标日期
                        target_date = current_date + timedelta(weeks=week_offset)
                        # 调整到指定的星期几
                        target_date = target_date - timedelta(days=target_date.weekday()) + timedelta(days=day-1)
                        target_dates.append(target_date.strftime("%Y-%m-%d"))
                
            elif mode == "biweekly":
                # 单双周：复制到隔周的指定星期几
                weekdays = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
                selected = sd.askstring("单双周复制", "选择星期几(1-7)\n1:周一 2:周二 ... 7:周日", 
                                      initialvalue=str(current_date.isoweekday()))
                
                if not selected:
                    return
                    
                try:
                    day = int(selected.strip())
                    if day < 1 or day > 7:
                        raise ValueError("星期值必须在1-7之间")
                except Exception as e:
                    messagebox.showerror("错误", f"无效的输入: {str(e)}")
                    return
                
                # 计算未来8周内隔周的指定星期几
                for week_offset in range(1, 9, 2):  # 隔周，共8周
                    # 计算目标日期
                    target_date = current_date + timedelta(weeks=week_offset)
                    # 调整到指定的星期几
                    target_date = target_date - timedelta(days=target_date.weekday()) + timedelta(days=day-1)
                    target_dates.append(target_date.strftime("%Y-%m-%d"))
                
            elif mode == "daily":
                # 按日期：复制到指定日期范围内的每一天
                start_date = sd.askstring("按日期复制", "开始日期(YYYY-MM-DD)", 
                                         initialvalue=self.selected_date)
                end_date = sd.askstring("按日期复制", "结束日期(YYYY-MM-DD)", 
                                       initialvalue=(current_date + timedelta(days=7)).strftime("%Y-%m-%d"))
                
                if not start_date or not end_date:
                    return
                    
                try:
                    start = datetime.strptime(start_date, "%Y-%m-%d")
                    end = datetime.strptime(end_date, "%Y-%m-%d")
                    
                    if start > end:
                        messagebox.showerror("错误", "开始日期不能晚于结束日期")
                        return
                        
                    # 生成日期范围内的所有日期
                    current = start
                    while current <= end:
                        target_dates.append(current.strftime("%Y-%m-%d"))
                        current += timedelta(days=1)
                except Exception as e:
                    messagebox.showerror("错误", f"日期格式错误: {str(e)}")
                    return
            
            # 将事件复制到所有目标日期
            count = 0
            for date_str in target_dates:
                # 创建新事件
                new_event = {
                    "time": time_val,
                    "task": task_val,
                    "completion": "待评价"  # 设置为待评价
                }
                
                if date_str not in self.events:
                    self.events[date_str] = []
                
                # 添加到事件列表
                self.events[date_str].append(new_event)
                
                # 按时间排序
                self.events[date_str] = sorted(
                    self.events[date_str], 
                    key=lambda x: self.time_to_minutes(x["time"])
                )
                
                count += 1
            
            # 更新日历和显示
            self.update_calendar()
            if self.context_row and self.context_col:
                self.show_events(self.context_row, self.context_col)
                
            self.modified = True
            
            # 只显示复制成功的消息
            messagebox.showinfo("成功", f"已复制事件到 {count} 个日期")
            
        except Exception as e:
            print(f"应用格式刷时出错: {str(e)}")
            messagebox.showerror("错误", f"应用格式刷时出错: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CalendarApp(root)
    root.mainloop()