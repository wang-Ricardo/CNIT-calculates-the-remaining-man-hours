#!/usr/bin/python
from datetime import datetime, date, time
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from dataclasses import dataclass
from typing import List, Tuple, Optional
import logging
import sqlite3
import json
from pathlib import Path
import requests
import os

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@dataclass
class TimeConfig:
    """时间配置类"""
    WORK_START = time(8, 30)
    WORK_LATE = time(9, 30)
    WORK_END = time(18, 0)
    WORK_END_FLEX = time(18, 30)
    STANDARD_WORK_MINUTES = 570  # 9.5小时
    OVERTIME_WORK_MINUTES = 600  # 10小时

@dataclass
class AttendanceResult:
    """考勤结果类"""
    name: str                      
    month: str                     
    overtime_hours: float
    missing_clockout_dates: List[date]
    late_count: int
    start_date: date
    end_date: date
    valid_days: int

class HolidayCalendar:
    _instance = None
    _initialized = False
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        if not HolidayCalendar._initialized:
            # 获取Windows系统的LocalAppData目录
            local_app_data = os.getenv('LOCALAPPDATA')
            if not local_app_data:
                local_app_data = os.path.expanduser('~\\AppData\\Local')
                
            # 创建应用专用目录
            app_dir = os.path.join(local_app_data, 'AttendanceSystem')
            os.makedirs(app_dir, exist_ok=True)
            
            # 设置数据库和缓存目录路径
            self.db_path = Path(app_dir) / 'holidays.db'
            self.cache_dir = Path(app_dir) / 'holiday_cache'
            self.cache_dir.mkdir(exist_ok=True)
            
            self.holidays = set()
            self.workdays = set()
            
            self._init_database()
            self._load_all_data()
            HolidayCalendar._initialized = True

    def _init_database(self):
        """初始化数据库"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 创建节假日表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS holidays (
                date TEXT PRIMARY KEY,
                type TEXT,
                description TEXT,
                year INTEGER
            )
        ''')
        
        conn.commit()
        conn.close()

    def _fetch_holiday_data(self, year: int) -> dict:
        """从API获取节假日数据"""
        url = f'http://timor.tech/api/holiday/year/{year}'
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            data = response.json()
            self._save_to_cache(year, data)  # 保存到缓存
            return data
        return None
    
    def is_workday(self, date_obj: date) -> bool:
        """判断是否为工作日"""
        if date_obj in self.holidays:
            return False
        if date_obj in self.workdays:
            return True
        return date_obj.weekday() < 5
    
    def _load_all_data(self):
        """加载所有数据"""
        current_year = datetime.now().year
        years_to_load = range(current_year - 1, current_year + 2)
        
        # 首先尝试从数据库加载
        for year in years_to_load:
            if self._load_from_database(year):
                continue
                
            # 如果数据库没有数据，尝试从缓存加载
            cache_data = self._load_from_cache(year)
            if cache_data:
                self._save_to_database(year, cache_data)
                self._load_from_database(year)

    def _save_to_database(self, year: int, data: dict):
        """保存数据到数据库"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            holiday_data = []
            for date_str, info in data.get('holiday', {}).items():
                try:
                    # 确保日期格式正确
                    if len(date_str.split('-')) == 2:
                        date_str = f"{year}-{date_str}"
                    datetime.strptime(date_str, '%Y-%m-%d')  # 验证日期格式
                    
                    holiday_data.append(
                        (date_str, 'holiday' if info['holiday'] else 'workday', 
                         info.get('name', ''), year)
                    )
                except ValueError:
                    logger.warning(f"跳过无效日期格式: {date_str}")
                    continue
            
            cursor.executemany(
                'INSERT OR REPLACE INTO holidays VALUES (?, ?, ?, ?)',
                holiday_data
            )
            conn.commit()
            
        except Exception as e:
            logger.error(f"保存数据到数据库失败: {e}")
            conn.rollback()
        finally:
            conn.close()

    def _load_from_database(self, year: int) -> bool:
        """从数据库加载特定年份的节假日数据"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 一次性获取所有数据
            cursor.execute("""
                SELECT date, type 
                FROM holidays 
                WHERE year = ? 
                AND (type = 'holiday' OR type = 'workday')
            """, (year,))
            
            results = cursor.fetchall()
            conn.close()
            
            if not results:
                return False
                
            # 批量处理数据
            for date_str, type_ in results:
                try:
                    # 确保日期字符串包含年份
                    if len(date_str.split('-')) == 2:
                        date_str = f"{year}-{date_str}"
                    
                    date_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
                    if type_ == 'holiday':
                        self.holidays.add(date_obj)
                    else:
                        self.workdays.add(date_obj)
                except ValueError as e:
                    logger.warning(f"跳过无效日期格式: {date_str}, 错误: {e}")
                    continue
            
            return True
            
        except Exception as e:
            logger.error(f"从数据库加载节假日数据失败: {e}")
            return False

    def force_update(self):
        """强制更新数据"""
        current_year = datetime.now().year
        years_to_update = range(current_year - 1, current_year + 2)
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            for year in years_to_update:
                api_data = self._fetch_holiday_data(year)
                if not api_data:
                    continue
                    
                # 批量插入数据
                holiday_data = [
                    (date_str, 'holiday' if info['holiday'] else 'workday', 
                     info.get('name', ''), year)
                    for date_str, info in api_data.get('holiday', {}).items()
                ]
                
                cursor.executemany(
                    'INSERT OR REPLACE INTO holidays VALUES (?, ?, ?, ?)',
                    holiday_data
                )
            
            conn.commit()
            
            # 重新加载数据
            self.holidays.clear()
            self.workdays.clear()
            for year in years_to_update:
                self._load_from_database(year)
                
        except Exception as e:
            logger.error(f"更新节假日数据失败: {e}")
            conn.rollback()
        finally:
            conn.close()

    def _save_to_cache(self, year: int, data: dict):
        """保存数据到缓存文件"""
        cache_file = self.cache_dir / f'holiday_{year}.json'
        try:
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False)
        except Exception as e:
            logger.error(f"保存缓存数据失败: {e}")

    def _load_from_cache(self, year: int) -> Optional[dict]:
        """从缓存文件加载数据"""
        cache_file = self.cache_dir / f'holiday_{year}.json'
        try:
            if cache_file.exists():
                with open(cache_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            logger.error(f"加载缓存数据失败: {e}")
        return None

class TimeParser:
    """时间解析类"""
    @staticmethod
    def parse_time(time_str: str) -> Optional[time]:
        """解析时间字符串"""
        try:
            return datetime.strptime(time_str, '%H:%M').time()
        except ValueError:
            return None

    @staticmethod
    def parse_next_day_time(time_str: str) -> int:
        """解析次日时间"""
        time_str = time_str.replace("次日", "")
        time_obj = datetime.strptime(time_str, '%H:%M')
        return time_obj.hour * 60 + time_obj.minute

    @staticmethod
    def compute_minutes(start_time: str, end_time: str) -> int:
        """计算时间差（分钟）"""
        start_dt = datetime.strptime(start_time, "%H:%M")
        end_dt = datetime.strptime(end_time, "%H:%M")
        diff = end_dt - start_dt
        return int(diff.total_seconds() / 60)

class AttendanceAnalyzer:
    """考勤分析类"""
    def __init__(self):
        self.time_config = TimeConfig()
        self.calendar = HolidayCalendar()
        self.time_parser = TimeParser()

    def analyze_attendance(self, file_path: str) -> AttendanceResult:
        """分析考勤数据"""
        try:
            df = pd.read_excel(file_path)
            selected_data = df.iloc[3:, [0, 8, 9, 10, 14]]
            selected_data = selected_data.iloc[::-1]

            try:
                name = df.iloc[3, 1]
                month_str = df.iloc[3, 0].split()[0]
                date_obj = datetime.strptime(month_str, '%Y/%m/%d')
                month = f"{date_obj.month:02d}"
            except:
                name = ""
                month = ""

            total_overtime = 0
            missing_clockout = []
            late_count = 0
            valid_days = 0

            first_date = datetime.strptime(
                selected_data.iloc[0, 0].split()[0], 
                '%Y/%m/%d'
            ).date()
            last_date = datetime.strptime(
                selected_data.iloc[-1, 0].split()[0], 
                '%Y/%m/%d'
            ).date()

            for _, row in selected_data.iterrows():
                date_str = row[0].split()[0]
                date_obj = datetime.strptime(date_str, '%Y/%m/%d').date()
                
                punch_count = row[3]        # 打卡次数
                start_time = row[1]         # 开始时间
                end_time = row[2]           # 结束时间
                
                # 使用日历判断是否为工作日（已经考虑了调休情况）
                is_workday = self.calendar.is_workday(date_obj)
                
                has_actual_punch = (pd.notna(start_time) and pd.notna(end_time) and 
                                str(start_time) != str(end_time))
                
                # 统计有效打卡天数（包括工作日和非工作日）
                if pd.notna(punch_count) and punch_count == "2次":
                    if is_workday or (not is_workday and has_actual_punch):
                        valid_days += 1
                        logger.debug(f"有效打卡日期: {date_str}, 是否工作日: {is_workday}")
                
                # 只计算工作日的加班时间
                overtime, is_late, missing = self._process_record(row)
                if is_workday:  # 只累计工作日的加班时间
                    total_overtime += overtime
                    late_count += is_late
                if missing:
                    missing_clockout.append(missing)

            missing_clockout.sort(reverse=True)

            return AttendanceResult(
                name=name,
                month=month,
                overtime_hours=total_overtime / 60,  # 只包含工作日的加班时间
                missing_clockout_dates=missing_clockout,
                late_count=late_count,
                start_date=first_date,
                end_date=last_date,
                valid_days=valid_days  # 包含所有有效打卡天数
            )

        except Exception as e:
            logger.error(f"分析考勤数据时出错: {e}")
            raise

    def _process_record(self, row) -> Tuple[float, bool, Optional[date]]:
        """处理单条考勤记录"""
        date_str = row[0].split()[0]
        date_obj = datetime.strptime(date_str, '%Y/%m/%d').date() #  将日期字符串转换为日期对象
        
        # 判断是否为工作日
        is_workday = self.calendar.is_workday(date_obj)

        start_time = row[1]
        end_time = str(row[2])
        work_status = row[4]
        punch_count = row[3]
        
        # 如果是当天的记录且开始时间等于结束时间，说明还在上班，跳过计算
        if end_time == start_time:
            return 0, False, None
        
        # 检查考勤状态和打卡次数
        if work_status != "正常" or (pd.notna(punch_count) and punch_count != "2次"):
            return 0, False, date_obj if date_obj != date.today() else None

            
        # 计算加班时间
        if pd.isna(start_time) or pd.isna(end_time):
            return 0, False, date_obj if date_obj != date.today() else None
            
        # 传入是否工作日的标志
        return self._calculate_overtime(start_time, end_time, is_workday), \
            self._is_late(start_time) if is_workday else False, \
            None

    def _is_before_time(self, time_str: str, target: time) -> bool:
        """判断是否早于某个时间点"""
        time_obj = self.time_parser.parse_time(time_str)
        return time_obj and time_obj <= target

    def _is_late(self, time_str: str) -> bool:
        """判断是否迟到"""
        time_obj = self.time_parser.parse_time(time_str)
        return time_obj and self.time_config.WORK_START < time_obj <= self.time_config.WORK_LATE

    def _calculate_overtime(self, start_time: str, end_time: str, is_workday: bool = True) -> float:
        """计算加班时间"""
        # 非工作日不计算剩余工时
        if not is_workday:
            return 0
            
        if "次日" in end_time:
            return self.time_parser.parse_next_day_time(end_time) + 1 +\
                self.time_parser.compute_minutes(start_time, "23:59") - \
                self.time_config.OVERTIME_WORK_MINUTES

        if self._is_before_time(end_time, self.time_config.WORK_END_FLEX):
            work_minutes = self.time_config.STANDARD_WORK_MINUTES
            end_time = "18:00"
        else:
            work_minutes = self.time_config.OVERTIME_WORK_MINUTES
            
        return self.time_parser.compute_minutes(start_time, end_time) - work_minutes

def main():
    root = tk.Tk()
    root.withdraw()
    
    try:
        file_path = filedialog.askopenfilename()
        if not file_path:
            return
            
        analyzer = AttendanceAnalyzer()
        result = analyzer.analyze_attendance(file_path)
        logger.info(f"姓名: {result.name}")
        logger.info(f"月份: {result.month}")
        logger.info(f"加班时长: {result.overtime_hours:.2f}小时")
        logger.info(f"有效打卡天数: {result.valid_days}天")  # 新增：输出有效打卡天数
        if result.missing_clockout_dates:
            logger.info(f"缺少打卡记录的日期: {result.missing_clockout_dates}")
        logger.info(f"迟到次数: {result.late_count}")
        logger.info(f"统计期间: {result.start_date} 至 {result.end_date}")
        
    except Exception as e:
        logger.error(f"程序执行出错: {e}")

if __name__ == "__main__":
    main()