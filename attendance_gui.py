#!/usr/bin/env python
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, messagebox
import win32ui
from typing import Optional
from dataclasses import dataclass
import logging
from Calculate_working_hours import AttendanceAnalyzer

logger = logging.getLogger(__name__)

@dataclass
class DisplayData:
    """显示数据类"""
    name: str = ''
    attendance_type: str = ''
    month: str = ''
    total_hours: float = 0.0
    non_flex_hours: float = 0.0
    valid_days: int = 0
    flex_count: int = 0
    start_date: str = ''
    end_date: str = ''

class CustomProgressBar(tk.Canvas):
    """自定义进度条，带刻度线和标记"""
    def __init__(self, master, width=350, height=60, **kwargs):  # 缩短进度条宽度
        super().__init__(master, width=width, height=height, **kwargs)
        self.configure(highlightthickness=0, bg='white')
        self.width = width
        self.height = height
        self.bar_height = 30
        self.value = 0
        
        # 进度条的刻度点（小时）
        self.scale_points = [22, 25, 28, 32, 35, 40, 45]
        # 进度条对应的天数值和位置
        self.day_markers = [
            {"value": 1, "start": 22, "end": 28, "position": 24},
            {"value": 2, "start": 28, "end": 35, "position": 30},
            {"value": 3, "start": 35, "end": 45, "position": 38},
            {"value": 4, "start": 45, "end": 50, "position": 48}
        ]
        
        self.create_base_elements()
        
    def create_base_elements(self):
        """创建基础元素"""
        # 绘制进度条边框 - 使用黑色边框
        self.create_rectangle(10, 10, self.width-10, 10+self.bar_height, 
                            outline='black', width=1, fill='white')
        
        # 绘制刻度线和标签
        scale_width = self.width - 20
        
        # 绘制天数标记（在进度条内部，错开竖线位置）
        for marker in self.day_markers:
            # 计算标记位置 - 错开竖线
            x_pos = 10 + (marker["position"] / 50) * scale_width
            
            # 绘制天数标签（在进度条内部）
            self.create_text(x_pos, 10 + self.bar_height/2, 
                           text=str(marker["value"]), fill='black', font=('SimHei', 10))
        
        # 绘制小时刻度线和标签（在进度条下方）
        for point in self.scale_points:
            # 计算刻度线位置
            x_pos = 10 + (point / 50) * scale_width
            
            # 绘制刻度线 - 从进度条内部延伸到下方，使用黑色
            self.create_line(x_pos, 10, x_pos, 10+self.bar_height, 
                           fill='black', width=1)
            
            # 绘制小时标签 - 在进度条下方，使用黑色
            self.create_text(x_pos, 10+self.bar_height+10, 
                           text=str(point), fill='black', font=('SimHei', 8))
    
    def update_value(self, hours):
        """更新进度条值"""
        self.delete("progress")
        
        if hours <= 0:
            return 0
            
        # 计算进度条宽度
        scale_width = self.width - 20
        bar_width = min((hours / 50) * scale_width, scale_width)
        
        # 根据小时数计算天数值
        day_value = 0
        if hours >= 45:
            day_value = 4
        elif hours >= 35:
            day_value = 3
        elif hours >= 28:
            day_value = 2
        elif hours >= 22:
            day_value = 1
            
        # 绘制进度条填充
        if bar_width > 0:
            self.create_rectangle(10, 10, 10 + bar_width, 10+self.bar_height, 
                                outline='', fill='#4a86e8', tags="progress")
            
        return day_value

class AttendanceGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.filename: Optional[str] = None
        self.analyzer = AttendanceAnalyzer()
        
        self._init_window()
        self._create_widgets()

    def _init_window(self):
        """初始化窗口设置"""
        self.root.title('考勤分析系统')
        
        # 设置窗口大小和位置
        window_width = 700
        window_height = 450
        x = (self.root.winfo_screenwidth() - window_width) // 2
        y = (self.root.winfo_screenheight() - window_height) // 2
        
        self.root.geometry(f'{window_width}x{window_height}+{x}+{y}')
        self.root.resizable(False, False)

    def _create_widgets(self):
        """创建窗口部件"""
        # 创建左右分区
        left_frame = ttk.Frame(self.root)
        right_frame = ttk.Frame(self.root)
        
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=5)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 创建标题
        ttk.Label(left_frame, text="考勤信息", font=('SimHei', 10)).pack(anchor=tk.W, pady=(0, 10))
        ttk.Label(right_frame, text="使用说明", font=('SimHei', 10)).pack(anchor=tk.W, pady=(0, 10))

        # 创建基本信息区域
        info_frame = ttk.Frame(left_frame)
        info_frame.pack(fill=tk.X, pady=5)

        # 修改为竖排布局 - 为每个项目创建单独的行
        # 姓名行
        name_row = ttk.Frame(info_frame)
        name_row.pack(fill=tk.X, pady=2)
        self.name_label = ttk.Label(name_row, text="姓名：", font=('SimHei', 10))
        self.name_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # 月份行
        month_row = ttk.Frame(info_frame)
        month_row.pack(fill=tk.X, pady=2)
        self.month_label = ttk.Label(month_row, text="月份：", font=('SimHei', 10))
        self.month_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # 统计期间行
        period_row = ttk.Frame(info_frame)
        period_row.pack(fill=tk.X, pady=2)
        self.period_label = ttk.Label(period_row, text="统计期间：", font=('SimHei', 10))
        self.period_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # 剩余工时行
        hours_row = ttk.Frame(info_frame)
        hours_row.pack(fill=tk.X, pady=2)
        self.hours_label = ttk.Label(hours_row, text="剩余工时：0.00", font=('SimHei', 10))
        self.hours_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # 弹性次数行
        flex_row = ttk.Frame(info_frame)
        flex_row.pack(fill=tk.X, pady=2)
        self.flex_label = ttk.Label(flex_row, text="弹性次数：0", font=('SimHei', 10))
        self.flex_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # 有效打卡天数行
        valid_days_row = ttk.Frame(info_frame)
        valid_days_row.pack(fill=tk.X, pady=2)
        self.valid_days_label = ttk.Label(valid_days_row, text="有效打卡天数：0", font=('SimHei', 10))
        self.valid_days_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # 创建中间占位区域（用于留出空间）
        spacer_frame = ttk.Frame(left_frame)
        spacer_frame.pack(fill=tk.BOTH, expand=True)
        
        # 进度条容器框架
        progress_container = ttk.Frame(left_frame)
        progress_container.pack(fill=tk.X, padx=5, pady=(5, 30))
        
        # 在进度条上方添加"可加班天数"标签 - 使用黑色字体
        ttk.Label(progress_container, text="可加班天数", font=('SimHei', 10)).pack(anchor=tk.W, pady=(0, 3))
        
        # 创建自定义进度条
        self.progress_frame = ttk.Frame(progress_container)
        self.progress_frame.pack(fill=tk.X, pady=2)
        
        self.custom_progress = CustomProgressBar(self.progress_frame, width=320, height=60)  # 进一步缩短宽度
        self.custom_progress.pack(side=tk.LEFT, expand=False)  # 设置expand=False防止拉伸
        
        # 创建显示具体可加班天数的标签
        self.days_label = ttk.Label(
            self.progress_frame, 
            text="0天", 
            font=('SimHei', 10),
            width=4
        )
        self.days_label.pack(side=tk.RIGHT, padx=(10, 0))

        # 创建说明文本 - 修复换行问题
        instruction_text = (
            "1. 智能分析特性:\n\n"
            "   √ 自动识别节假日和调休日期\n\n"
            "   √ 智能修正打卡时间误差\n\n"
            "   √ 工时计算精确到分钟\n\n"
            "   √ 异常打卡自动标记\n\n"
            "2. 计算规则:\n\n"
            "   - 标准工时: 8小时/天\n\n"
            "   - 加班工时: 当日总工时减标准工时\n\n"
            "   - 晚餐时间: 18:00 - 18:30不计入工时\n\n"
            "3. 数据更新:\n\n"
            "   - 可自动同步国家法定假日\n\n"
            "4. 工时统计截止时间:\n\n"
            "   - 加班邮件提交加班前一天\n\n"
        )
        self.instruction_text = tk.Text(
            right_frame,
            wrap=tk.WORD,
            width=35,  # 增加宽度以容纳更多文本
            height=17,
            font=('SimHei', 10),
            bg=self.root.cget('bg'),
            relief='flat',
            padx=5,
            pady=5
        )
        self.instruction_text.insert('1.0', instruction_text)
        self.instruction_text.config(state='disabled')
        self.instruction_text.pack(fill=tk.BOTH, expand=True)

        # 创建按钮框架
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(side=tk.BOTTOM, pady=20)

        # 创建选择文件按钮
        self.select_file_btn = ttk.Button(
            button_frame,
            text="请选择考勤文件",
            command=self.open_file,
            width=20
        )
        self.select_file_btn.pack(side=tk.LEFT, padx=5)

        # 创建更新缓存按钮
        self.update_cache_btn = ttk.Button(
            button_frame,
            text="更新节假日数据",
            command=self.update_cache_data,
            width=20
        )
        self.update_cache_btn.pack(side=tk.LEFT, padx=5)

    def update_cache_data(self):
        """更新缓存数据"""
        try:
            self.update_cache_btn.config(state='disabled')
            self.update_cache_btn.config(text="正在更新...")
            
            # 更新数据
            self.analyzer.calendar.force_update()
            
            messagebox.showinfo("成功", "节假日数据更新成功！")
        except Exception as e:
            logger.error(f"更新节假日数据失败: {e}")
            messagebox.showerror("错误", f"更新节假日数据失败: {e}")
        finally:
            self.update_cache_btn.config(state='normal')
            self.update_cache_btn.config(text="更新节假日数据")
            
    def update_display(self, result):
        """更新显示数据"""
        self.name_label.config(text=f"姓名：{result.name}")
        self.month_label.config(text=f"月份：{result.month}")
        self.period_label.config(text=f"统计期间：{result.start_date} 至 {result.end_date}")
        self.hours_label.config(text=f"剩余工时：{result.overtime_hours:.2f}")
        self.flex_label.config(text=f"弹性次数：{result.late_count}")
        self.valid_days_label.config(text=f"有效打卡天数：{result.valid_days}")
        
        # 更新自定义进度条
        day_value = self.custom_progress.update_value(result.overtime_hours)
        self.days_label.config(text=f"{day_value}天")
        
        # 根据可加班天数设置不同的颜色
        if day_value >= 3:
            self.days_label.config(foreground="green")
        elif day_value >= 1:
            self.days_label.config(foreground="blue")
        else:
            self.days_label.config(foreground="gray")
        
        if result.missing_clockout_dates:
            messagebox.showwarning(
                "警告", 
                f"以下日期缺少打卡记录：\n{', '.join(str(d) for d in result.missing_clockout_dates)}"
            )

    def open_file(self):
        """打开文件对话框"""
        try:
            dlg = win32ui.CreateFileDialog(1)
            dlg.SetOFNInitialDir(r'C:\Users\Administrator\Desktop')
            
            if dlg.DoModal() == 1:
                self.filename = dlg.GetPathName()
                if self.filename:
                    result = self.analyzer.analyze_attendance(self.filename)
                    self.update_display(result)
        except Exception as e:
            logger.error(f"处理文件时出错: {e}")
            messagebox.showerror("错误", f"处理文件时出错: {e}")

def main():
    root = tk.Tk()
    app = AttendanceGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()