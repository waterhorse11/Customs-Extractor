import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
import queue
import logging
import sys

# 导入我们后端逻辑的工厂类
from ExtractorFactory import ExtractorFactory

class TkinterLogHandler(logging.Handler):
    """一个将日志记录发送到线程安全队列的处理器。"""
    def __init__(self, queue):
        super().__init__()
        self.queue = queue

    def emit(self, record):
        msg = self.format(record)
        self.queue.put(msg)

class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Customs Form Extractor")
        self.geometry("650x550")

        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        self.thread = None

        # --- 数据模型 ---
        self.template_options = {
            'import': ("LSS", "HLS", "SNP", "OLC", "TianShi"),
            'export': ("TianShi", "HLS")
        }

        # --- UI组件 ---
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        
        # 配置main_frame的列权重，使内容能够水平拉伸
        main_frame.columnconfigure(0, weight=1)

        # PDF文件选择 (row=0)
        pdf_frame = ttk.LabelFrame(main_frame, text="PDF File")
        pdf_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.pdf_path_var = tk.StringVar()
        self.pdf_path_entry = ttk.Entry(pdf_frame, textvariable=self.pdf_path_var)
        self.pdf_browse_button = ttk.Button(pdf_frame, text="Browse...", command=self.browse_pdf)
        self.pdf_path_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.pdf_browse_button.grid(row=0, column=1, padx=5, pady=5)
        pdf_frame.columnconfigure(0, weight=1)

        # 输出目录选择 (row=1)
        output_frame = ttk.LabelFrame(main_frame, text="Output Directory")
        output_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.output_path_var = tk.StringVar()
        self.output_path_entry = ttk.Entry(output_frame, textvariable=self.output_path_var)
        self.output_browse_button = ttk.Button(output_frame, text="Browse...", command=self.browse_output_dir)
        self.output_path_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.output_browse_button.grid(row=0, column=1, padx=5, pady=5)
        output_frame.columnconfigure(0, weight=1)
        
        # 选项 (row=2)
        options_frame = ttk.LabelFrame(main_frame, text="Extraction Options")
        options_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5, padx=5)

        # 模板类型 (import/export)
        self.template_type_label = ttk.Label(options_frame, text="Template Type:")
        self.template_type_label.pack(side=tk.LEFT, padx=(5, 5))
        self.template_type_var = tk.StringVar()
        self.template_type_combo = ttk.Combobox(options_frame, textvariable=self.template_type_var, state="readonly", width=10)
        self.template_type_combo['values'] = list(self.template_options.keys())
        self.template_type_combo.pack(side=tk.LEFT, padx=(0, 15))
        self.template_type_combo.bind("<<ComboboxSelected>>", self._update_template_options)

        # PDF 模板
        self.template_label = ttk.Label(options_frame, text="PDF Template:")
        self.template_label.pack(side=tk.LEFT, padx=(5, 5))
        self.template_var = tk.StringVar()
        self.template_combo = ttk.Combobox(options_frame, textvariable=self.template_var, state="readonly", width=10)
        self.template_combo.pack(side=tk.LEFT, padx=(0, 15))
        
        # 初始化默认选项
        self.template_type_combo.current(0)
        self._update_template_options()

        # 开始按钮 (row=3)
        actions_frame = ttk.Frame(main_frame)
        actions_frame.grid(row=3, column=0, columnspan=2, sticky=tk.EW, pady=10)
        actions_frame.columnconfigure(0, weight=1)  # actions_frame 内部列权重1
        
        self.start_button = ttk.Button(actions_frame, text="Start Extraction", command=self.run_extraction_in_thread)
        self.start_button.grid(row=0, column=0)

        # 进度条 (row=4)
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.progress_bar = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(side=tk.LEFT, expand=True, fill='x')
        self.progress_percent_label = ttk.Label(progress_frame, text="0%")
        self.progress_percent_label.pack(side=tk.RIGHT, padx=5)
        
        # 日志显示区域 (row=5)
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log")
        log_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.log_display = scrolledtext.ScrolledText(log_frame, state='disabled', wrap=tk.WORD, height=10)
        self.log_display.pack(expand=True, fill='both', padx=5, pady=5)
        
        # 配置main_frame的行权重，使日志区域能够垂直拉伸
        main_frame.rowconfigure(5, weight=1)

        self.process_queues()

    def _update_template_options(self, event=None):
        """当模板类型改变时，更新PDF模板的下拉选项。"""
        selected_type = self.template_type_var.get()
        options = self.template_options.get(selected_type, [])
        self.template_combo['values'] = options
        if options:
            self.template_combo.current(0)
        else:
            self.template_var.set('')

    def browse_pdf(self):
        file_name = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])
        if file_name:
            self.pdf_path_var.set(file_name)
            if not self.output_path_var.get():
                pdf_dir = os.path.dirname(file_name)
                self.output_path_var.set(os.path.join(pdf_dir, "output"))
                if not os.path.exists(self.output_path_var.get()):
                    os.makedirs(self.output_path_var.get())

    def browse_output_dir(self):
        dir_name = filedialog.askdirectory(title="Select Output Directory")
        if dir_name:
            self.output_path_var.set(dir_name)
    
    def update_log(self, message):
        self.log_display.config(state='normal')
        self.log_display.insert(tk.END, message + '\n')
        self.log_display.see(tk.END)
        self.log_display.config(state='disabled')
        
    def process_queues(self):
        """同时处理日志和进度条队列。"""
        try:
            log_message = self.log_queue.get_nowait()
            self.update_log(log_message)
        except queue.Empty:
            pass

        try:
            progress_value = self.progress_queue.get_nowait()
            self.progress_bar['value'] = progress_value
            self.progress_percent_label.config(text=f"{progress_value}%")
        except queue.Empty:
            pass
        finally:
            self.after(100, self.process_queues)

    def _extraction_task(self, pdf_path, output_dir, template_type, type_name):
        try:
            self.after(0, self.update_log, f"Processing: {os.path.basename(pdf_path)}")
            self.after(0, self.update_log, f"Using Template: {template_type} (Type: {type_name})")
            
            extractor = ExtractorFactory.create_extractor(
                template_type=template_type,
                pdf_path=pdf_path,
                output_dir=output_dir,
                type=type_name,
            )
            
            # 将进度队列传递给提取器
            if hasattr(extractor, 'ocr_parser'):
                extractor.ocr_parser.progress_queue = self.progress_queue

            # 配置日志处理器
            log_handler = TkinterLogHandler(self.log_queue)
            log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            
            # 清理旧的处理器并添加新的
            extractor.logger.handlers.clear()
            extractor.logger.addHandler(log_handler)
            extractor.logger.setLevel(logging.INFO)

            extractor.extract_items()
            
            final_message = f"Processing completed!\nResults saved to: {output_dir}"
            self.after(0, lambda: messagebox.showinfo("Completed", final_message))
            self.after(0, self.update_log, f"{os.path.basename(pdf_path)} Processing completed!")
        
        except Exception as e:
            error_message = f"Error: {e}"
            self.after(0, lambda: messagebox.showerror("Error", error_message))
            
        finally:
            self.after(0, lambda: self.start_button.config(state="normal"))
            self.after(0, lambda: self.progress_bar.config(value=0))
            self.after(0, lambda: self.progress_percent_label.config(text="0%"))

    def run_extraction_in_thread(self):
        pdf_path = self.pdf_path_var.get()
        output_dir = self.output_path_var.get()
        template_type = self.template_var.get()
        type_name = self.template_type_var.get()

        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showwarning("Error", "Please select a valid PDF file.")
            return
            
        if not output_dir:
            messagebox.showwarning("Error", "Please select an output directory.")
            return
            
        if not template_type:
            messagebox.showwarning("Error", "Please select a PDF template.")
            return

        self.start_button.config(state="disabled")
        self.progress_bar['value'] = 0
        self.progress_percent_label.config(text="0%")
        self.log_display.config(state='normal')
        self.log_display.delete('1.0', tk.END)
        self.log_display.config(state='disabled')
        
        self.thread = threading.Thread(
            target=self._extraction_task,
            args=(pdf_path, output_dir, template_type, type_name)
        )
        self.thread.daemon = True
        self.thread.start()
if __name__ == '__main__':
    import multiprocessing
    multiprocessing.freeze_support()
    
    app = MainWindow()
    app.mainloop()

