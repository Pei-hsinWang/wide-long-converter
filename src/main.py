"""
通用宽面板转长面板转换工具
版本：2.0.0
功能：支持用户自定义 ID 列、值列、正则提取规则
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import threading
import re
import json
from datetime import datetime
from typing import List, Dict, Optional


class WideToLongConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # 窗口配置
        self.title("📊 通用宽面板转长面板工具 v2.0")
        self.geometry("900x750")
        self.resizable(True, True)
        
        # 状态变量
        self.input_path = None
        self.output_path = None
        self.df = None
        self.column_info = {}
        self.is_processing = False
        
        # 配置变量
        self.id_columns = []
        self.value_columns = []
        self.variable_pattern = r"(\d+)"  # 默认提取数字
        self.variable_name = "变量"
        self.value_name = "值"
        self.extract_column_name = "提取值"
        
        # 创建界面
        self.create_widgets()
    
    def create_widgets(self):
        # ============ 顶部：文件选择区域 ============
        file_frame = ctk.CTkFrame(self)
        file_frame.pack(pady=10, padx=20, fill="x")
        
        # 输入文件
        input_subframe = ctk.CTkFrame(file_frame, fg_color="transparent")
        input_subframe.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(input_subframe, text="📂 输入文件:", width=100, anchor="w").pack(side="left")
        self.input_btn = ctk.CTkButton(input_subframe, text="选择文件", command=self.select_input, width=100)
        self.input_btn.pack(side="left", padx=5)
        self.input_label = ctk.CTkLabel(input_subframe, text="未选择", text_color="gray")
        self.input_label.pack(side="left", padx=5, fill="x", expand=True)
        
        # 输出文件
        output_subframe = ctk.CTkFrame(file_frame, fg_color="transparent")
        output_subframe.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(output_subframe, text="💾 输出文件:", width=100, anchor="w").pack(side="left")
        self.output_btn = ctk.CTkButton(output_subframe, text="保存位置", command=self.select_output, width=100)
        self.output_btn.pack(side="left", padx=5)
        self.output_label = ctk.CTkLabel(output_subframe, text="未选择", text_color="gray")
        self.output_label.pack(side="left", padx=5, fill="x", expand=True)
        
        # ============ 中部：列配置区域 ============
        config_frame = ctk.CTkFrame(self)
        config_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # 配置说明
        config_header = ctk.CTkFrame(config_frame, fg_color="transparent")
        config_header.pack(fill="x", padx=10, pady=(10, 5))
        
        ctk.CTkLabel(
            config_header, 
            text="📋 列配置 (选择 ID 列和值列)", 
            font=ctk.CTkFont(weight="bold", size=14)
        ).pack(side="left")
        
        self.auto_detect_btn = ctk.CTkButton(
            config_header, 
            text="🔍 自动检测", 
            command=self.auto_detect_columns,
            width=100,
            height=30
        )
        self.auto_detect_btn.pack(side="right", padx=5)
        
        # 列列表区域（带复选框）
        list_frame = ctk.CTkFrame(config_frame)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 创建可滚动框架
        self.scroll_frame = ctk.CTkScrollableFrame(list_frame, height=200)
        self.scroll_frame.pack(fill="both", expand=True)
        
        # 列标题
        header_frame = ctk.CTkFrame(self.scroll_frame, fg_color="#2B2B2B")
        header_frame.pack(fill="x")
        
        ctk.CTkLabel(header_frame, text="列名", width=200, anchor="w").pack(side="left", padx=10, pady=5)
        ctk.CTkLabel(header_frame, text="类型", width=100, anchor="w").pack(side="left", padx=10, pady=5)
        ctk.CTkLabel(header_frame, text="前 3 个值", width=200, anchor="w").pack(side="left", padx=10, pady=5)
        ctk.CTkLabel(header_frame, text="作为 ID 列", width=100, anchor="w").pack(side="left", padx=10, pady=5)
        ctk.CTkLabel(header_frame, text="作为值列", width=100, anchor="w").pack(side="left", padx=10, pady=5)
        
        # 列配置容器
        self.column_vars_frame = ctk.CTkFrame(self.scroll_frame)
        self.column_vars_frame.pack(fill="both", expand=True)
        
        self.column_checkboxes = {}  # 存储每列的复选框变量
        
        # ============ 下部：正则和输出配置 ============
        pattern_frame = ctk.CTkFrame(self)
        pattern_frame.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(
            pattern_frame, 
            text="🔧 提取规则配置", 
            font=ctk.CTkFont(weight="bold", size=14)
        ).pack(anchor="w", padx=10, pady=(10, 5))
        
        # 正则表达式
        pattern_subframe = ctk.CTkFrame(pattern_frame, fg_color="transparent")
        pattern_subframe.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(pattern_subframe, text="正则表达式:", width=120, anchor="w").pack(side="left")
        self.pattern_entry = ctk.CTkEntry(pattern_subframe, width=200, placeholder_text="如：(\\d+) 提取数字")
        self.pattern_entry.pack(side="left", padx=5)
        self.pattern_entry.insert(0, r"(\d+)")
        
        ctk.CTkLabel(pattern_subframe, text="示例：PM2000 → 2000", text_color="gray").pack(side="left", padx=10)
        
        # 输出列名
        name_subframe = ctk.CTkFrame(pattern_frame, fg_color="transparent")
        name_subframe.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(name_subframe, text="变量列名:", width=120, anchor="w").pack(side="left")
        self.variable_name_entry = ctk.CTkEntry(name_subframe, width=150)
        self.variable_name_entry.pack(side="left", padx=5)
        self.variable_name_entry.insert(0, "年份")
        
        ctk.CTkLabel(name_subframe, text="值列名:", width=80, anchor="w").pack(side="left", padx=10)
        self.value_name_entry = ctk.CTkEntry(name_subframe, width=150)
        self.value_name_entry.pack(side="left", padx=5)
        self.value_name_entry.insert(0, "PM2.5")
        
        ctk.CTkLabel(name_subframe, text="提取值列名:", width=100, anchor="w").pack(side="left", padx=10)
        self.extract_name_entry = ctk.CTkEntry(name_subframe, width=150)
        self.extract_name_entry.pack(side="left", padx=5)
        self.extract_name_entry.insert(0, "提取值")
        
        # ============ 底部：操作按钮 ============
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=15, padx=20, fill="x")
        
        self.save_config_btn = ctk.CTkButton(
            button_frame, 
            text="💾 保存配置模板", 
            command=self.save_config,
            width=130,
            fg_color="#555555"
        )
        self.save_config_btn.pack(side="left", padx=5)
        
        self.load_config_btn = ctk.CTkButton(
            button_frame, 
            text="📂 加载配置模板", 
            command=self.load_config,
            width=130,
            fg_color="#555555"
        )
        self.load_config_btn.pack(side="left", padx=5)
        
        self.preview_btn = ctk.CTkButton(
            button_frame, 
            text="👁️ 预览转换结果", 
            command=self.preview_conversion,
            width=130,
            fg_color="#FFA500"
        )
        self.preview_btn.pack(side="left", padx=5)
        
        self.convert_btn = ctk.CTkButton(
            button_frame, 
            text="🚀 开始转换", 
            command=self.start_conversion,
            height=40,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#2CC985",
            hover_color="#25A56E",
            state="disabled"
        )
        self.convert_btn.pack(side="right", padx=5)
        
        # ============ 状态栏 ============
        self.status_bar = ctk.CTkFrame(self, fg_color="#1a1a1a")
        self.status_bar.pack(fill="x", side="bottom")
        
        self.status_label = ctk.CTkLabel(
            self.status_bar, 
            text="就绪 - 请选择 Excel 文件开始", 
            text_color="gray"
        )
        self.status_label.pack(pady=5, padx=10)
        
        # ============ 日志区域（可折叠） ============
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.pack(pady=10, padx=20, fill="x")
        
        log_header = ctk.CTkFrame(self.log_frame, fg_color="transparent")
        log_header.pack(fill="x")
        
        ctk.CTkLabel(
            log_header, 
            text="📝 运行日志", 
            font=ctk.CTkFont(weight="bold")
        ).pack(side="left", padx=10, pady=5)
        
        self.log_text = ctk.CTkTextbox(self.log_frame, height=100)
        self.log_text.pack(fill="x", padx=10, pady=(0, 10))
    
    def log(self, message):
        """添加日志"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
    
    def update_status(self, message, color="white"):
        """更新状态栏"""
        self.status_label.configure(text=message, text_color=color)
    
    def select_input(self):
        """选择输入文件"""
        file_path = filedialog.askopenfilename(
            title="选择输入 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.input_path = file_path
            self.input_label.configure(text=file_path, text_color="#2CC985")
            self.log(f"已选择输入文件：{os.path.basename(file_path)}")
            self.load_file_preview()
            self.check_ready()
    
    def load_file_preview(self):
        """加载文件预览"""
        try:
            self.df = pd.read_excel(self.input_path, nrows=100)  # 只读取前 100 行用于预览
            self.log(f"文件加载成功：{len(self.df)} 行预览，{len(self.df.columns)} 列")
            self.populate_column_list()
            self.update_status("文件已加载 - 请配置列选项", "#2CC985")
        except Exception as e:
            self.log(f"❌ 文件加载失败：{e}")
            self.update_status("文件加载失败", "red")
    
    def populate_column_list(self):
        """填充列列表"""
        # 清空现有内容
        for widget in self.column_vars_frame.winfo_children():
            widget.destroy()
        self.column_checkboxes = {}
        
        # 分析每列信息
        for col in self.df.columns:
            col_data = self.df[col]
            
            # 检测列类型
            if col_data.dtype == 'object':
                col_type = "文本"
            elif col_data.dtype in ['int64', 'float64']:
                col_type = "数值"
            else:
                col_type = str(col_data.dtype)
            
            # 获取前 3 个值
            sample_values = col_data.head(3).tolist()
            sample_str = ", ".join([str(v) for v in sample_values])[:50]
            
            # 创建行
            row_frame = ctk.CTkFrame(self.column_vars_frame, fg_color="transparent")
            row_frame.pack(fill="x", pady=2)
            
            # 列名
            ctk.CTkLabel(row_frame, text=str(col), width=200, anchor="w").pack(side="left", padx=10)
            
            # 类型
            ctk.CTkLabel(row_frame, text=col_type, width=100, text_color="gray").pack(side="left", padx=10)
            
            # 样本值
            ctk.CTkLabel(row_frame, text=sample_str, width=200, text_color="gray").pack(side="left", padx=10)
            
            # ID 列复选框
            id_var = tk.BooleanVar(value=False)
            id_check = ctk.CTkCheckBox(row_frame, text="", variable=id_var, width=20)
            id_check.pack(side="left", padx=10)
            
            # 值列复选框
            value_var = tk.BooleanVar(value=False)
            value_check = ctk.CTkCheckBox(row_frame, text="", variable=value_var, width=20)
            value_check.pack(side="left", padx=10)
            
            self.column_checkboxes[col] = {
                'id_var': id_var,
                'value_var': value_var,
                'col_type': col_type
            }
    
    def auto_detect_columns(self):
        """自动检测 ID 列和值列"""
        if self.df is None:
            messagebox.showwarning("警告", "请先加载 Excel 文件")
            return
        
        self.log("开始自动检测列类型...")
        
        # 重置所有选择
        for col, vars_dict in self.column_checkboxes.items():
            vars_dict['id_var'].set(False)
            vars_dict['value_var'].set(False)
        
        # 检测模式 1: PM2000, PM2001 格式
        pm_pattern = re.compile(r'^[A-Za-z]+(\d{4})$')
        pm_columns = []
        for col in self.df.columns:
            if pm_pattern.match(str(col)):
                pm_columns.append(col)
        
        # 检测模式 2: 2000, 2001 纯数字格式
        numeric_columns = []
        for col in self.df.columns:
            try:
                int(str(col))
                numeric_columns.append(col)
            except:
                pass
        
        # 检测模式 3: Jan, Feb 月份格式
        month_pattern = re.compile(r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)', re.I)
        month_columns = []
        for col in self.df.columns:
            if month_pattern.match(str(col)):
                month_columns.append(col)
        
        # 选择检测到的值列
        value_cols_detected = pm_columns or numeric_columns or month_columns
        
        if value_cols_detected:
            self.log(f"✓ 检测到 {len(value_cols_detected)} 个可能的值列")
            for col in value_cols_detected:
                if col in self.column_checkboxes:
                    self.column_checkboxes[col]['value_var'].set(True)
            
            # 其余列作为 ID 列
            for col in self.df.columns:
                if col not in value_cols_detected and col in self.column_checkboxes:
                    self.column_checkboxes[col]['id_var'].set(True)
            
            self.log(f"✓ 检测到 {len(self.df.columns) - len(value_cols_detected)} 个 ID 列")
            
            # 自动设置正则
            if pm_columns:
                self.pattern_entry.delete(0, 'end')
                self.pattern_entry.insert(0, r"(\d+)")
                self.log("自动设置正则：提取数字 (适用于 PM2000 格式)")
            elif numeric_columns:
                self.pattern_entry.delete(0, 'end')
                self.pattern_entry.insert(0, r"(.*)")
                self.log("自动设置正则：提取全部 (适用于纯数字列名)")
            elif month_columns:
                self.pattern_entry.delete(0, 'end')
                self.pattern_entry.insert(0, r"([A-Za-z]+)")
                self.log("自动设置正则：提取字母 (适用于月份格式)")
            
            self.update_status("自动检测完成 - 请确认配置", "#2CC985")
        else:
            self.log("⚠️ 未检测到明显的值列模式，请手动选择")
            self.update_status("请手动选择 ID 列和值列", "orange")
    
    def select_output(self):
        """选择输出文件"""
        file_path = filedialog.asksaveasfilename(
            title="选择保存位置",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
            initialfile=f"long_format_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if file_path:
            self.output_path = file_path
            self.output_label.configure(text=file_path, text_color="#2CC985")
            self.log(f"已选择输出文件：{os.path.basename(file_path)}")
            self.check_ready()
    
    def check_ready(self):
        """检查是否可以开始转换"""
        if self.input_path and self.output_path and self.df is not None:
            self.convert_btn.configure(state="normal")
        else:
            self.convert_btn.configure(state="disabled")
    
    def get_current_config(self):
        """获取当前配置"""
        id_columns = []
        value_columns = []
        
        for col, vars_dict in self.column_checkboxes.items():
            if vars_dict['id_var'].get():
                id_columns.append(col)
            if vars_dict['value_var'].get():
                value_columns.append(col)
        
        return {
            'id_columns': id_columns,
            'value_columns': value_columns,
            'variable_pattern': self.pattern_entry.get(),
            'variable_name': self.variable_name_entry.get(),
            'value_name': self.value_name_entry.get(),
            'extract_column_name': self.extract_name_entry.get()
        }
    
    def save_config(self):
        """保存配置模板"""
        config = self.get_current_config()
        
        file_path = filedialog.asksaveasfilename(
            title="保存配置模板",
            defaultextension=".json",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")],
            initialfile="converter_config.json"
        )
        
        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            self.log(f"配置已保存：{file_path}")
            messagebox.showinfo("成功", "配置模板已保存！")
    
    def load_config(self):
        """加载配置模板"""
        file_path = filedialog.askopenfilename(
            title="加载配置模板",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # 应用配置
                for col, vars_dict in self.column_checkboxes.items():
                    vars_dict['id_var'].set(col in config.get('id_columns', []))
                    vars_dict['value_var'].set(col in config.get('value_columns', []))
                
                self.pattern_entry.delete(0, 'end')
                self.pattern_entry.insert(0, config.get('variable_pattern', r'(\d+)'))
                
                self.variable_name_entry.delete(0, 'end')
                self.variable_name_entry.insert(0, config.get('variable_name', '年份'))
                
                self.value_name_entry.delete(0, 'end')
                self.value_name_entry.insert(0, config.get('value_name', '值'))
                
                self.extract_name_entry.delete(0, 'end')
                self.extract_name_entry.insert(0, config.get('extract_column_name', '提取值'))
                
                self.log(f"配置已加载：{file_path}")
                messagebox.showinfo("成功", "配置模板已加载！")
            except Exception as e:
                self.log(f"❌ 加载配置失败：{e}")
                messagebox.showerror("错误", f"加载配置失败:\n{e}")
    
    def preview_conversion(self):
        """预览转换结果"""
        try:
            config = self.get_current_config()
            
            if not config['id_columns']:
                messagebox.showwarning("警告", "请至少选择一列作为 ID 列")
                return
            
            if not config['value_columns']:
                messagebox.showwarning("警告", "请至少选择一列作为值列")
                return
            
            # 读取完整数据
            df = pd.read_excel(self.input_path)
            
            # 执行转换
            long_df = pd.melt(
                df,
                id_vars=config['id_columns'],
                value_vars=config['value_columns'],
                var_name=config['variable_name'],
                value_name=config['value_name']
            )
            
            # 提取值
            try:
                pattern = config['variable_pattern']
                long_df[config['extract_column_name']] = long_df[config['variable_name']].str.extract(pattern)
            except Exception as e:
                self.log(f"⚠️ 正则提取失败：{e}")
                long_df[config['extract_column_name']] = long_df[config['variable_name']]
            
            # 显示预览
            preview_window = ctk.CTkToplevel(self)
            preview_window.title("预览转换结果")
            preview_window.geometry("600x400")
            
            text_box = ctk.CTkTextbox(preview_window)
            text_box.pack(fill="both", expand=True, padx=10, pady=10)
            
            preview_text = f"转换预览 (前 20 行):\n\n"
            preview_text += long_df.head(20).to_string()
            preview_text += f"\n\n总行数：{len(long_df)}"
            preview_text += f"\n总列数：{len(long_df.columns)}"
            preview_text += f"\n\n输出列：{list(long_df.columns)}"
            
            text_box.insert("1.0", preview_text)
            
            self.log("预览窗口已打开")
            
        except Exception as e:
            self.log(f"❌ 预览失败：{e}")
            messagebox.showerror("错误", f"预览失败:\n{e}")
    
    def start_conversion(self):
        """开始转换"""
        if self.is_processing:
            return
        
        config = self.get_current_config()
        
        if not config['id_columns']:
            messagebox.showwarning("警告", "请至少选择一列作为 ID 列")
            return
        
        if not config['value_columns']:
            messagebox.showwarning("警告", "请至少选择一列作为值列")
            return
        
        self.is_processing = True
        self.convert_btn.configure(state="disabled", text="⏳ 处理中...")
        self.log("=" * 50)
        self.log("开始转换任务...")
        self.log(f"ID 列：{config['id_columns']}")
        self.log(f"值列：{config['value_columns']} ({len(config['value_columns'])} 列)")
        
        thread = threading.Thread(target=self.run_conversion, args=(config,), daemon=True)
        thread.start()
    
    def run_conversion(self, config):
        """执行数据转换"""
        start_time = datetime.now()
        
        try:
            # 1. 读取完整数据
            self.update_status("正在读取数据...", "white")
            self.log("正在读取完整 Excel 文件...")
            
            df = pd.read_excel(self.input_path)
            self.log(f"✓ 读取成功：{len(df)} 行，{len(df.columns)} 列")
            
            # 2. 执行 melt
            self.update_status("正在转换数据...", "white")
            self.log("正在执行宽转长转换...")
            
            long_df = pd.melt(
                df,
                id_vars=config['id_columns'],
                value_vars=config['value_columns'],
                var_name=config['variable_name'],
                value_name=config['value_name']
            )
            
            self.log(f"✓ 转换完成：{len(long_df)} 行")
            
            # 3. 提取值
            self.update_status("正在提取变量值...", "white")
            self.log(f"使用正则 '{config['variable_pattern']}' 提取...")
            
            try:
                pattern = config['variable_pattern']
                extracted = long_df[config['variable_name']].str.extract(pattern)
                long_df[config['extract_column_name']] = extracted.iloc[:, 0] if extracted.shape[1] > 0 else long_df[config['variable_name']]
            except Exception as e:
                self.log(f"⚠️ 正则提取失败，使用原始值：{e}")
                long_df[config['extract_column_name']] = long_df[config['variable_name']]
            
            # 4. 排序
            self.log("正在排序...")
            if config['id_columns']:
                long_df = long_df.sort_values(by=config['id_columns'] + [config['extract_column_name']]).reset_index(drop=True)
            
            # 5. 保存
            self.update_status("正在保存文件...", "white")
            self.log(f"正在保存到：{self.output_path}")
            
            long_df.to_excel(self.output_path, index=False, sheet_name='LongFormat')
            
            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()
            
            self.log("=" * 50)
            self.log(f"✅ 转换成功!")
            self.log(f"  输出行数：{len(long_df)}")
            self.log(f"  输出列：{list(long_df.columns)}")
            self.log(f"  耗时：{duration:.2f}秒")
            
            self.update_status(f"转换完成！{len(long_df)}行 - 耗时{duration:.2f}秒", "#2CC985")
            
            self.after(0, lambda: messagebox.showinfo(
                "✅ 成功",
                f"转换完成!\n\n"
                f"📊 输出行数：{len(long_df)}\n"
                f"📋 输出列：{list(long_df.columns)}\n"
                f"⏱️  耗时：{duration:.2f}秒"
            ))
            
        except Exception as e:
            self.log(f"❌ 错误：{e}")
            self.update_status("转换失败", "red")
            self.after(0, lambda: messagebox.showerror("❌ 错误", f"转换失败:\n\n{e}"))
        
        finally:
            self.is_processing = False
            self.after(0, lambda: self.convert_btn.configure(state="normal", text="🚀 开始转换"))
    
    def update_status(self, message, color):
        """更新状态栏（线程安全）"""
        self.after(0, lambda: self.status_label.configure(text=message, text_color=color))


def main():
    app = WideToLongConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()