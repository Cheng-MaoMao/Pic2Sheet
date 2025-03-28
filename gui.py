import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from main import ImageAnalyzer, ExcelHandler
import threading
import queue
import json

class ImageAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.config_path = "config.json"
        self.load_config()  # 加载配置
        self.root.title("图片转表格工具")
        self.root.geometry("600x400")
        self.root.minsize(600, 550)  # 设置最小窗口大小
        
        # 配置根窗口的网格权重
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置主框架的网格权重
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)  # 状态文本框可扩展
        
        # API设置区域
        api_frame = ttk.LabelFrame(main_frame, text="API设置", padding="5")
        api_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        api_frame.grid_columnconfigure(1, weight=1)

        # 提供商选择
        ttk.Label(api_frame, text="提供商:").grid(row=0, column=0, sticky=tk.W)
        self.provider = tk.StringVar(value="阿里")
        provider_combobox = ttk.Combobox(api_frame, textvariable=self.provider, 
                                        values=["阿里", "火山引擎"], 
                                        state="readonly", width=10)
        provider_combobox.grid(row=0, column=1, sticky=tk.W, padx=5)
        provider_combobox.bind('<<ComboboxSelected>>', self.on_provider_change)

        # API Key输入
        ttk.Label(api_frame, text="API Key:").grid(row=1, column=0, sticky=tk.W)
        self.api_key = tk.StringVar(value="")
        api_key_entry = ttk.Entry(api_frame, textvariable=self.api_key, width=50)
        api_key_entry.grid(row=1, column=1, padx=5)
        api_key_entry.bind('<FocusOut>', self.on_config_change)  # 添加失去焦点事件

        # 模型选择
        ttk.Label(api_frame, text="模型名称:").grid(row=2, column=0, sticky=tk.W)
        self.model_name = tk.StringVar(value="qwen-vl-max-latest")
        self.model_entry = ttk.Entry(api_frame, textvariable=self.model_name, width=50)
        self.model_entry.grid(row=2, column=1, padx=5)
        self.model_entry.bind('<FocusOut>', self.on_config_change)  # 添加失去焦点事件
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件设置", padding="5")
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        file_frame.grid_columnconfigure(1, weight=1)  # Entry可扩展

        # 添加处理模式选择
        mode_frame = ttk.Frame(file_frame)
        mode_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.process_mode = tk.StringVar(value="single")
        ttk.Radiobutton(mode_frame, text="单文件处理", 
                        variable=self.process_mode, 
                        value="single",
                        command=self.on_mode_change).grid(row=0, column=0, padx=5)
        ttk.Radiobutton(mode_frame, text="批量处理", 
                        variable=self.process_mode, 
                        value="batch",
                        command=self.on_mode_change).grid(row=0, column=1, padx=5)
        ttk.Radiobutton(mode_frame, text="合并处理", 
                        variable=self.process_mode, 
                        value="merge",
                        command=self.on_mode_change).grid(row=0, column=2, padx=5)

        # 图片列表框
        self.image_listbox = tk.Listbox(file_frame, height=5, selectmode=tk.EXTENDED)
        self.image_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # 图片列表操作按钮框
        list_btn_frame = ttk.Frame(file_frame)
        list_btn_frame.grid(row=1, column=2, sticky=(tk.N, tk.S))
        
        ttk.Button(list_btn_frame, text="添加", 
                   command=self.add_images).grid(row=0, column=0, pady=2)
        ttk.Button(list_btn_frame, text="删除", 
                   command=self.remove_selected).grid(row=1, column=0, pady=2)
        ttk.Button(list_btn_frame, text="清空", 
                   command=self.clear_list).grid(row=2, column=0, pady=2)

        # Excel保存路径
        save_frame = ttk.Frame(file_frame)
        save_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        save_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Label(save_frame, text="保存路径:").grid(row=0, column=0, sticky=tk.W)
        self.save_path = tk.StringVar()
        ttk.Entry(save_frame, textvariable=self.save_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(save_frame, text="浏览", 
                   command=self.select_save_path).grid(row=0, column=2)
        
        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.E, tk.W), pady=5)
        button_frame.grid_columnconfigure(0, weight=1)  # 使按钮居中
        
        ttk.Button(button_frame, text="开始分析", command=self.start_analysis).grid(
            row=0, column=0, pady=5
        )
        
        # 添加进度条
        self.progress_var = tk.StringVar(value="就绪")
        self.progress_label = ttk.Label(button_frame, textvariable=self.progress_var)
        self.progress_label.grid(row=1, column=0, pady=5)
        
        # 创建进度条但初始不显示
        self.progress = ttk.Progressbar(button_frame, mode='indeterminate')
        self.progress.grid_remove()  # 初始化时隐藏进度条
        
        # 用于存储异步操作的结果
        self.result_queue = queue.Queue()
        
        # 状态显示
        self.status_text = tk.Text(main_frame, height=10, width=70)
        self.status_text.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=3, column=2, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        # 添加图片路径变量（用于单文件模式）
        self.image_path = tk.StringVar()
        
    def handle_existing_file(self, filepath):
        """处理已存在的文件"""
        if os.path.exists(filepath):
            answer = messagebox.askquestion(
                "文件已存在",
                "文件已存在，是否覆盖？\n选择'否'将自动重命名文件",
                icon='warning'
            )
            
            if answer == 'yes':
                return filepath
            else:
                # 自动重命名文件
                base_path = os.path.splitext(filepath)[0]
                ext = os.path.splitext(filepath)[1]
                counter = 1
                while os.path.exists(f"{base_path}_{counter}{ext}"):
                    counter += 1
                return f"{base_path}_{counter}{ext}"
        return filepath

    def select_image(self):
        filename = filedialog.askopenfilename(
            title="选择图片",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp")]
        )
        if filename:
            self.image_path.set(filename)
            # 自动设置默认保存路径并处理文件已存在的情况
            default_save = os.path.splitext(filename)[0] + ".xlsx"
            default_save = self.handle_existing_file(default_save)
            self.save_path.set(default_save)
    
    def select_save_path(self):
        filename = filedialog.asksaveasfilename(
            title="选择保存位置",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        if filename:
            # 处理文件已存在的情况
            filename = self.handle_existing_file(filename)
            self.save_path.set(filename)
    
    def load_config(self):
        """加载配置文件"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                    # 读取当前提供商的设置
                    provider = self.provider.get()
                    if provider in self.config:
                        self.api_key.set(self.config[provider]["api_key"])
                        self.model_name.set(self.config[provider]["model"])
            else:
                # 默认配置
                self.config = {
                    "阿里": {
                        "api_key": "",
                        "model": "qwen-vl-max-latest",
                        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1"
                    },
                    "火山引擎": {
                        "api_key": "",
                        "model": "doubao-1-5-vision-pro-32k-250115",
                        "base_url": "https://ark.cn-beijing.volces.com/api/v3"
                    }
                }
                self.save_config()
        except Exception as e:
            print(f"加载配置文件失败: {str(e)}")
            self.config = {}
    
    def save_config(self):
        """保存配置文件"""
        try:
            # 更新配置
            provider = self.provider.get()
            if provider not in self.config:
                self.config[provider] = {}
            self.config[provider]["api_key"] = self.api_key.get()
            self.config[provider]["model"] = self.model_name.get()
            
            # 保存到文件
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
                
            print(f"配置已保存: {provider} - API Key: {self.api_key.get()}, Model: {self.model_name.get()}")
        except Exception as e:
            print(f"保存配置文件失败: {str(e)}")
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
    
    def on_provider_change(self, event=None):
        """处理提供商切换"""
        provider = self.provider.get()
        if provider in self.config:
            # 加载对应提供商的配置
            self.api_key.set(self.config[provider]["api_key"])
            self.model_name.set(self.config[provider]["model"])
        else:
            # 新提供商，使用默认值
            self.api_key.set("")
            self.model_name.set("qwen-vl-max-latest" if provider == "阿里" else "doubao-1-5-vision-pro-32k-250115")
        # 保存配置
        self.save_config()

    def on_mode_change(self):
        """处理模式切换"""
        mode = self.process_mode.get()
        if mode == "single":
            # 清空列表
            self.clear_list()
            # 清空保存路径
            self.save_path.set("")

    def process_image(self):
        """在后台线程中处理图片"""
        try:
            # 检查 API Key
            api_key = self.api_key.get()
            if not api_key:
                raise ValueError("请先设置 API Key")
                
            # 从列表框获取图片路径（单文件模式）
            if self.image_listbox.size() > 0:
                image_path = self.image_listbox.get(0)
                
                # 保存当前配置
                self.save_config()
                
                # 创建分析器并处理图片
                analyzer = ImageAnalyzer(
                    provider=self.provider.get(),
                    api_key=api_key
                )
                result = analyzer.analyze_image(image_path)
                if not result:
                    raise ValueError("图片分析失败")
                self.result_queue.put(("success", result))
            else:
                raise ValueError("请先选择要分析的图片")
        except Exception as e:
            self.result_queue.put(("error", str(e)))

    def check_result(self):
        """检查后台处理是否完成"""
        try:
            status, result = self.result_queue.get_nowait()
            
            # 停止并隐藏进度条
            self.progress.stop()
            self.progress.grid_remove()
            
            if status == "success":
                mode = self.process_mode.get()
                if mode == "merge":
                    # 合并模式
                    if result is None:  # 成功完成所有处理
                        messagebox.showinfo("成功", "所有图片处理完成并保存")
                    else:
                        self.status_text.insert(tk.END, "合并处理失败\n")
                elif mode == "batch":
                    # 批量处理模式
                    if result is None:  # 成功完成所有处理
                        messagebox.showinfo("成功", "批量处理完成")
                    else:
                        self.status_text.insert(tk.END, "批量处理失败\n")
                else:
                    # 单文件模式
                    if result:  # result 包含处理结果
                        excel_handler = ExcelHandler(self.save_path.get())
                        if excel_handler.write_data(result):
                            self.status_text.insert(tk.END, f"分析完成！表格已保存至: {self.save_path.get()}\n")
                            messagebox.showinfo("成功", "分析完成并保存")
                        else:
                            self.status_text.insert(tk.END, "保存表格时发生错误\n")
                    else:
                        self.status_text.insert(tk.END, "分析失败\n")
            else:
                self.status_text.insert(tk.END, f"发生错误: {result}\n")
                messagebox.showerror("错误", result)
            
            self.progress_var.set("就绪")
            
        except queue.Empty:
            # 如果队列为空，说明处理还未完成，继续检查
            self.root.after(100, self.check_result)
    
    def start_analysis(self):
        mode = self.process_mode.get()
        
        if mode in ["batch", "merge"]:
            if not self.get_image_list():
                messagebox.showerror("错误", "请添加要处理的图片")
                return
        else:  # single mode
            if not self.image_path.get():
                messagebox.showerror("错误", "请选择要分析的图片")
                return
        
        if not self.save_path.get():
            messagebox.showerror("错误", "请选择保存位置")
            return
        
        # 显示并启动进度条
        self.progress_var.set("正在处理图片...")
        self.progress.grid(row=2, column=0, sticky=(tk.E, tk.W), pady=5)
        self.progress.start(10)
        
        # 清空状态文本
        self.status_text.delete(1.0, tk.END)
        self.status_text.insert(tk.END, "开始处理...\n")
        
        # 在后台线程中处理
        thread = threading.Thread(
            target=self.process_images if mode in ["batch", "merge"] else self.process_image
        )
        thread.daemon = True
        thread.start()
        
        # 开始检查结果
        self.check_result()

    def add_images(self):
        """添加图片到列表"""
        mode = self.process_mode.get()
        
        if mode == "single":
            # 单文件模式
            filename = filedialog.askopenfilename(
                title="选择图片",
                filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp")]
            )
            if filename:
                # 清空现有列表
                self.clear_list()
                # 添加新文件
                self.image_listbox.insert(tk.END, filename)
                self.image_path.set(filename)
                # 设置默认保存路径
                default_save = os.path.splitext(filename)[0] + ".xlsx"
                default_save = self.handle_existing_file(default_save)
                self.save_path.set(default_save)
        else:
            # 批量模式
            filenames = filedialog.askopenfilenames(
                title="选择图片",
                filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp")]
            )
            if filenames:
                for filename in filenames:
                    if filename not in self.get_image_list():
                        self.image_listbox.insert(tk.END, filename)
                
                # 自动设置默认保存路径
                if mode == "merge":
                    default_save = os.path.join(os.path.dirname(filenames[0]), "合并结果.xlsx")
                else:
                    default_save = os.path.dirname(filenames[0])
                self.save_path.set(default_save)

    def remove_selected(self):
        """删除选中的图片"""
        selected = self.image_listbox.curselection()
        for i in reversed(selected):
            self.image_listbox.delete(i)

    def clear_list(self):
        """清空图片列表"""
        self.image_listbox.delete(0, tk.END)

    def get_image_list(self):
        """获取图片列表"""
        return [self.image_listbox.get(i) for i in range(self.image_listbox.size())]

    def process_images(self):
        """批量处理图片"""
        try:
            # 检查 API Key
            api_key = self.api_key.get()
            if not api_key:
                raise ValueError("请先设置 API Key")
                
            image_list = self.get_image_list()
            mode = self.process_mode.get()
            
            # 创建分析器
            analyzer = ImageAnalyzer(
                provider=self.provider.get(),
                api_key=api_key
            )
            
            success = True
            
            if mode == "merge":
                # 合并处理模式 - 一次性处理所有图片
                self.status_text.insert(tk.END, "开始分析所有图片...\n")
                result = analyzer.analyze_images_batch(image_list)
                
                if result:
                    try:
                        # 提取 JSON 部分
                        import re
                        json_match = re.search(r'\[\s*\[.*?\]\s*\]', result, re.DOTALL)
                        if json_match:
                            json_str = json_match.group()
                            # 解析返回的JSON数据
                            data = json.loads(json_str)
                            excel_handler = ExcelHandler(self.save_path.get())
                            if excel_handler.write_data(json_str):
                                self.status_text.insert(tk.END, f"合并处理完成！表格已保存至: {self.save_path.get()}\n")
                                success = True
                            else:
                                success = False
                                self.status_text.insert(tk.END, "保存合并数据失败\n")
                        else:
                            self.status_text.insert(tk.END, "未找到有效的JSON数据\n")
                            success = False
                    except (json.JSONDecodeError, Exception) as e:
                        self.status_text.insert(tk.END, f"处理返回数据失败: {str(e)}\n")
                        success = False
                else:
                    self.status_text.insert(tk.END, "分析图片失败\n")
                    success = False
                    
            else:
                # 批量处理模式 - 逐个处理图片
                for image_path in image_list:
                    save_path = os.path.join(
                        self.save_path.get(),
                        os.path.splitext(os.path.basename(image_path))[0] + ".xlsx"
                    )
                    
                    result = analyzer.analyze_image(image_path)
                    if result:
                        excel_handler = ExcelHandler(save_path)
                        if excel_handler.write_data(result):
                            self.status_text.insert(tk.END, f"处理完成: {save_path}\n")
                        else:
                            self.status_text.insert(tk.END, f"处理失败: {image_path}\n")
                            success = False
                    else:
                        self.status_text.insert(tk.END, f"处理失败: {image_path}\n")
                        success = False
                        
                if success:
                    self.status_text.insert(tk.END, "批量处理完成！\n")
                else:
                    self.status_text.insert(tk.END, "批量处理完成，但部分文件处理失败\n")
            
            # 根据整体处理结果发送成功或失败状态
            if success:
                self.result_queue.put(("success", None))
            else:
                self.result_queue.put(("error", "部分文件处理失败"))
                
        except Exception as e:
            self.result_queue.put(("error", str(e)))

    def on_config_change(self, event=None):
        """处理配置变更"""
        self.save_config()

def main():
    root = tk.Tk()
    app = ImageAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()