from flask import Flask, request, jsonify, send_file
import os
import json
import numpy as np
from datetime import datetime, timedelta
from dotenv import load_dotenv
from openai import OpenAI
import re
import pandas as pd
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from flask_httpauth import HTTPBasicAuth
from filelock import FileLock
import traceback
import time  # 添加缺失的time模块导入

# 加载 .env 文件
load_dotenv()

app = Flask(__name__)

# 安全配置
auth = HTTPBasicAuth()
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=["1000 per day", "100 per hour"],
    storage_uri="memory://",
)

# 存储用户日程和反馈的数据结构
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB

# 获取脚本所在的文件夹路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# 构建 Excel 文件的完整路径
EXCEL_FILE_PATH = os.path.join(SCRIPT_DIR, "记录.xlsx")

# 认证配置
@auth.verify_password
def verify_password(username, password):
    return username == os.getenv("API_USERNAME", "admin") and \
           password == os.getenv("API_PASSWORD", "password")

class ModelIntegrator:
    def __init__(self, api_key, base_url="https://api.siliconflow.cn/v1"):
        self.client = OpenAI(
            api_key=api_key,
            base_url=base_url
        )

    def chat(self, messages, model="deepseek-ai/DeepSeek-V3", temperature=0.7, top_p=1.0, max_tokens=3000):
        response = self.client.chat.completions.create(
            model=model,
            messages=messages,
            temperature=temperature,
            top_p=top_p,
            max_tokens=max_tokens,
            stream=False
        )
        # 返回完整内容
        return response.choices[0].message.content

    def chat_stream(self, messages, model="deepseek-ai/DeepSeek-V3", temperature=0.7, top_p=1.0, max_tokens=3000):
        return self.client.chat.completions.create(
            model=model,
            messages=messages,
            temperature=temperature,
            top_p=top_p,
            max_tokens=max_tokens,
            stream=True
        )

class LLMAPI:
    def __init__(self, api_key, model="deepseek-ai/DeepSeek-V3"):
        self.integrator = ModelIntegrator(api_key)
        self.model = model

    def generate_response(self, prompt, temperature=0.7, top_p=1.0, max_tokens=3000):  
        try:
            print(f"调用LLM API - 模型: {self.model}, 温度: {temperature}, top_p: {top_p}, max_tokens: {max_tokens}")
            print(f"LLM输入内容: {prompt[:]}...")  # 只打印前500个字符
            
            response = self.integrator.chat(
                messages=[{"role": "user", "content": prompt}],
                model=self.model,
                temperature=temperature,
                top_p=top_p,
                max_tokens=max_tokens
            )
            
            print(f"LLM返回内容: {response[:]}...")  # 只打印前500个字符
            return response
        except Exception as e:
            return f"生成响应时发生错误: {str(e)}"

    def generate_response_stream(self, prompt, temperature=0.7, top_p=1.0, max_tokens=512):
        try:
            print(f"调用LLM流式API - 模型: {self.model}, 温度: {temperature}, top_p: {top_p}, max_tokens: {max_tokens}")
            print(f"LLM输入内容: {prompt[:]}...")
            
            gen = self.integrator.chat_stream(
                messages=[{"role": "user", "content": prompt}],
                model=self.model,
                temperature=temperature,
                top_p=top_p,
                max_tokens=max_tokens
            )
            
            for chunk in gen:
                # 打印流式返回的内容
                if hasattr(chunk, 'choices') and chunk.choices and hasattr(chunk.choices[0], 'delta'):
                    content = chunk.choices[0].delta.content or ""
                    print(f"LLM流式返回内容: {content[:]}...")  # 只打印前100个字符
                yield chunk
        except Exception as e:
            yield f"生成响应时发生错误: {str(e)}"

def get_current_week_id():
    """获取当前周ID (YYYY-WW格式，周从周一开始)"""
    return datetime.now().strftime("%Y-%W")

def get_next_week_id():
    """获取下一周ID"""
    next_week = datetime.now() + timedelta(weeks=1)
    return next_week.strftime("%Y-%W")

def parse_excel_date(date_str):
    """将Excel中的日期字符串解析为标准日期格式"""
    try:
        # 处理pandas日期类型
        if isinstance(date_str, pd.Timestamp):
            return date_str.strftime('%Y-%m-%d')
        
        # 处理datetime对象
        if isinstance(date_str, datetime):
            return date_str.strftime('%Y-%m-%d')
        
        # 处理字符串格式的日期
        if isinstance(date_str, str):
            # 尝试常见的日期格式
            for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d-%m-%Y', '%Y/%m/%d', '%d/%m/%Y', '%Y.%m.%d'):
                try:
                    return datetime.strptime(date_str, fmt).strftime('%Y-%m-%d')
                except ValueError:
                    continue
            
            # 尝试"年.月.日"格式
            if re.match(r"\d{4}\.\d{1,2}\.\d{1,2}", date_str):
                parts = date_str.split('.')
                if len(parts) == 3:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])
                    return f"{year}-{month:02d}-{day:02d}"
            
            # 尝试Excel数字日期格式
            try:
                date_num = float(date_str)
                if date_num > 0:
                    base_date = datetime(1899, 12, 30)
                    return (base_date + timedelta(days=date_num)).strftime('%Y-%m-%d')
            except ValueError:
                pass
                
            # 尝试提取数字部分
            match = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', date_str)
            if match:
                year, month, day = match.groups()
                return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
            
        return None
    except Exception as e:
        print(f"解析日期失败: {date_str}, 错误: {str(e)}")
        return None
    
def parse_excel_schedule(file_path):
    """使用与caption.py相同的方法解析Excel文件"""
    try:
        df = pd.read_excel(file_path, dtype=str)
        print(f"成功读取Excel文件，共{len(df)}行数据")
        
        # 打印数据前几行，检查数据格式
        print("数据前几行内容:")
        print(df.head().to_csv(sep='\t', na_rep='nan'))
        
        # 检查必要列
        required_columns = ["日期", "时间", "任务", "完成度"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            return None, f"Excel文件缺少必要列: {', '.join(missing_columns)}"
        
        # 按日期分组活动
        schedule = {}
        
        # 逐行处理数据
        for index, row in df.iterrows():
            # 解析日期 - 使用与caption.py相同的逻辑
            date_str = str(row["日期"])
            
            # 尝试多种日期格式
            parsed_date = None
            # 尝试"年.月.日"格式
            if re.match(r"\d{4}\.\d{1,2}\.\d{1,2}", date_str):
                parts = date_str.split('.')
                if len(parts) == 3:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])
                    parsed_date = f"{year}-{month:02d}-{day:02d}"
            
            # 处理pandas日期类型
            if not parsed_date and isinstance(date_str, pd.Timestamp):
                parsed_date = date_str.strftime("%Y-%m-%d")
            
            # 处理datetime对象
            if not parsed_date and isinstance(date_str, datetime):
                parsed_date = date_str.strftime("%Y-%m-%d")
            
            # 尝试其他常见格式
            if not parsed_date:
                for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S"):
                    try:
                        dt = datetime.strptime(date_str, fmt)
                        parsed_date = dt.strftime("%Y-%m-%d")
                        break
                    except ValueError:
                        continue
            
            # 尝试Excel序列号日期
            if not parsed_date:
                try:
                    if isinstance(date_str, (int, float)):
                        # Excel日期是从1900-01-01开始的天数
                        base_date = datetime(1900, 1, 1)
                        parsed_date = (base_date + timedelta(days=float(date_str) - 2).strftime("%Y-%m-%d"))
                except:
                    pass
            
            if not parsed_date:
                print(f"跳过无法解析的日期: {date_str} (行 {index+2})")
                continue
            
            # 标准化时间 - 使用与caption.py相同的逻辑
            time_str = str(row["时间"])
            # 尝试统一时间分隔符
            time_str = time_str.replace("：", ":")  # 替换中文冒号
            time_str = time_str.replace("—", "-")   # 替换中文破折号
            time_str = time_str.replace("~", "-")   # 替换波浪号
            
            # 处理时间段
            if "-" in time_str:
                parts = time_str.split("-")
                if len(parts) == 2:
                    start = normalize_single_time(parts[0].strip())
                    end = normalize_single_time(parts[1].strip())
                    time_str = f"{start} - {end}"
            else:
                # 处理单个时间点
                time_str = normalize_single_time(time_str.strip())
            
            # 提取活动类型（任务）
            task = str(row["任务"]) if pd.notna(row["任务"]) else ""
            
            # 提取完成度
            completion = str(row["完成度"]) if pd.notna(row["完成度"]) else "未开始"
            
            # 创建活动（删除持续时间字段）
            activity = {
                "type": task,
                "time": time_str,
                "completion": completion
            }
            
            # 按日期分组
            if parsed_date not in schedule:
                schedule[parsed_date] = {"activities": []}
            
            schedule[parsed_date]["activities"].append(activity)
        
        # 对每天的活动按时间排序
        for date in schedule:
            schedule[date]["activities"] = sorted(
                schedule[date]["activities"],
                key=lambda x: time_to_minutes(x["time"])
            )
        
        print(f"成功从Excel加载 {len(df)} 条活动记录")
        return schedule, None
        
    except Exception as e:
        print(f"解析Excel日程出错: {str(e)}")
        print(traceback.format_exc())
        return None, f"解析Excel出错: {str(e)}"

def parse_excel_feedback(file_path):
    """解析反馈Excel文件"""
    try:
        df = pd.read_excel(file_path, dtype=str)
        print(f"成功读取反馈Excel文件，共{len(df)}行数据")
        
        # 打印数据前几行，检查数据格式
        print("反馈数据前几行内容:")
        print(df.head().to_csv(sep='\t', na_rep='nan'))
        
        # 检查必要列
        required_columns = ["日期", "评分"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            return None, f"反馈Excel文件缺少必要列: {', '.join(missing_columns)}"
        
        # 按日期分组反馈
        feedback = {}
        
        # 逐行处理数据
        for index, row in df.iterrows():
            # 解析日期
            date_str = str(row["日期"])
            parsed_date = parse_excel_date(date_str)
            if not parsed_date:
                print(f"跳过无法解析的日期: {date_str} (行 {index+2})")
                continue
            
            # 提取评分
            try:
                rating = float(row["评分"])
            except:
                rating = 3.0  # 默认评分
                
            # 提取评论
            comments = str(row.get("评论", "")) if "评论" in df.columns and pd.notna(row.get("评论")) else ""
            
            # 创建反馈
            feedback[parsed_date] = {
                "rating": rating,
                "comments": comments
            }
        
        print(f"成功从Excel加载 {len(feedback)} 条反馈记录")
        return feedback, None
    except Exception as e:
        print(f"解析反馈Excel出错: {str(e)}")
        print(traceback.format_exc())
        return None, f"解析反馈Excel出错: {str(e)}"
def save_schedule_to_excel(schedule_data, file_path=EXCEL_FILE_PATH):
    """将日程数据保存到Excel文件（使用caption.py相同的格式）"""
    try:
        # 创建数据框
        rows = []
        for date, day_data in schedule_data.items():
            for activity in day_data.get("activities", []):
                rows.append({
                    "日期": date,
                    "时间": activity.get("time", "00:00"),
                    "任务": activity.get("type", ""),
                    "完成度": activity.get("completion", "待评价")
                })
        
        df = pd.DataFrame(rows)
        
        # 确保列顺序正确
        if not df.empty:
            df = df[["日期", "时间", "任务", "完成度"]]
        
        # 写入Excel
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='Schedule', index=False)
        
        print(f"成功将日程写入Excel: {file_path}")
        return True, None
    except Exception as e:
        print(f"写入Excel时出错: {str(e)}")
        return False, str(e)

def validate_excel_export(schedule_data, file_path):
    """验证Excel导出是否正确"""
    try:
        # 读取刚刚导出的Excel文件
        df = pd.read_excel(file_path, dtype=str)
        
        # 检查必要列
        required_columns = {"日期", "时间", "任务", "完成度"}
        if not required_columns.issubset(set(df.columns)):
            return False, f"缺少必要列: {', '.join(required_columns - set(df.columns))}"
        
        # 构建导出的数据结构用于比较
        exported_schedule = {}
        for _, row in df.iterrows():
            date = str(row.get("日期", "")).strip()
            if not date:
                continue
                
            time_str = normalize_single_time(str(row.get("时间", "")))
            task = str(row.get("任务", "")).strip()
            completion = str(row.get("完成度", "待评价")).strip()
            
            if date not in exported_schedule:
                exported_schedule[date] = {"activities": []}
                
            exported_schedule[date]["activities"].append({
                "type": task,
                "time": time_str,
                "completion": completion
            })
        
        # 对导出的数据按时间排序
        for date in exported_schedule:
            exported_schedule[date]["activities"] = sorted(
                exported_schedule[date]["activities"],
                key=lambda x: time_to_minutes(x["time"])
            )
        
        # 比较原数据和导出的数据
        for date, day_data in schedule_data.items():
            if date not in exported_schedule:
                return False, f"导出的日程中缺少日期: {date}"
                
            original_activities = day_data.get("activities", [])
            exported_activities = exported_schedule[date].get("activities", [])
            
            if len(original_activities) != len(exported_activities):
                return False, f"日期 {date} 的活动数量不匹配，预期: {len(original_activities)}，实际: {len(exported_activities)}"
                
            for i, (orig, exp) in enumerate(zip(original_activities, exported_activities)):
                if orig.get("type") != exp.get("type"):
                    return False, f"日期 {date} 的活动 {i+1} 类型不匹配，预期: {orig.get('type')}，实际: {exp.get('type')}"
                    
                if orig.get("time") != exp.get("time"):
                    return False, f"日期 {date} 的活动 {i+1} 时间不匹配，预期: {orig.get('time')}，实际: {exp.get('time')}"
                    
                if orig.get("completion") != exp.get("completion"):
                    return False, f"日期 {date} 的活动 {i+1} 完成度不匹配，预期: {orig.get('completion')}，实际: {exp.get('completion')}"
        
        return True, None
    except Exception as e:
        print(f"验证Excel导出时出错: {str(e)}")
        return False, str(e)
# 时间标准化辅助方法
def normalize_single_time(time_str):
    """标准化单个时间点格式（与caption.py相同）"""
    # 尝试解析时间格式
    if re.match(r"\d{1,2}:\d{2}", time_str):
        # 已经是标准格式
        return time_str
    
    # 尝试添加分钟部分
    if re.match(r"\d{1,2}$", time_str):
        return f"{time_str}:00"
    
    # 其他格式直接返回
    return time_str

def time_to_minutes(time_str):
    """将时间字符串转换为分钟数用于排序（与caption.py相同）"""
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

def normalize_single_time(time_str):
    """标准化单个时间点格式（与caption.py相同）"""
    # 尝试解析时间格式
    if re.match(r"\d{1,2}:\d{2}", time_str):
        # 已经是标准格式
        return time_str
    
    # 尝试添加分钟部分
    if re.match(r"\d{1,2}$", time_str):
        return f"{time_str}:00"
    
    # 其他格式直接返回
    return time_str

def time_to_minutes(time_str):
    """将时间字符串转换为分钟数用于排序（与caption.py相同）"""
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
    
if __name__ == '__main__':
    # 验证必要的环境变量
    required_env_vars = ["DEEPSEEK_API_KEY", "API_USERNAME", "API_PASSWORD"]
    missing_vars = [var for var in required_env_vars if not os.getenv(var)]
    
    if missing_vars:
        print(f"错误: 缺少必要的环境变量: {', '.join(missing_vars)}")
        print("请在.env文件中设置这些变量")
    else:
        app.run(debug=True)