import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from task_manager import TaskManager

class AppGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PPT批量加水印保护工具 (Windows专版)")
        self.geometry("600x620")
        
        self.files = []
        self.watermark_path = ""
        self.output_dir = ""
        
        self.task_manager = TaskManager(self.update_progress, self.on_finish)
        
        self.create_widgets()
        
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 1. PPT Files Selection
        ttk.Label(main_frame, text="1. 选择需要处理的 PPT 文件 (支持多选):").pack(anchor=tk.W, pady=(0,5))
        
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.file_listbox = tk.Listbox(file_frame, height=5)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(file_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=scrollbar.set)
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 15))
        ttk.Button(btn_frame, text="添加 PPT 文件", command=self.add_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="清空列表", command=self.clear_files).pack(side=tk.LEFT)
        
        # 2. Watermark Selection
        ttk.Label(main_frame, text="2. 选择水印图片 (建议使用透明背景 PNG):").pack(anchor=tk.W, pady=(0,5))
        wm_frame = ttk.Frame(main_frame)
        wm_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.wm_var = tk.StringVar()
        ttk.Entry(wm_frame, textvariable=self.wm_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(wm_frame, text="浏览...", command=self.select_watermark).pack(side=tk.RIGHT)
        
        # 3. Watermark Scale
        ttk.Label(main_frame, text="3. 水印缩放比例 (相对于屏幕高度的百分比):").pack(anchor=tk.W, pady=(0,5))
        scale_frame = ttk.Frame(main_frame)
        scale_frame.pack(fill=tk.X, pady=(0, 15))
        
        def on_scale_drag(val):
            # Snap to 1% steps when dragging
            self.scale_var.set(f"{round(float(val))}")
            
        self.scale_var = tk.StringVar(value="15")
        scale_slider = ttk.Scale(scale_frame, from_=5, to=50, variable=self.scale_var, orient=tk.HORIZONTAL, command=on_scale_drag)
        scale_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Use Entry to allow custom input with decimals
        scale_entry = ttk.Entry(scale_frame, textvariable=self.scale_var, width=6)
        scale_entry.pack(side=tk.LEFT)
        ttk.Label(scale_frame, text="%").pack(side=tk.LEFT)
        
        # 4. Resolution and Quality
        ttk.Label(main_frame, text="4. 导出分辨率与压缩质量 (决定最终文件大小):").pack(anchor=tk.W, pady=(0,5))
        rq_frame = ttk.Frame(main_frame)
        rq_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(rq_frame, text="分辨率:").pack(side=tk.LEFT)
        self.res_var = tk.StringVar(value="1080p")
        ttk.Combobox(rq_frame, textvariable=self.res_var, values=["1080p", "720p"], state="readonly", width=8).pack(side=tk.LEFT, padx=(5, 20))
        
        ttk.Label(rq_frame, text="JPG图片质量:").pack(side=tk.LEFT)
        
        def on_quality_drag(val):
            # Snap to integer (1% steps)
            self.quality_var.set(f"{round(float(val))}")
            
        self.quality_var = tk.StringVar(value="80")
        q_slider = ttk.Scale(rq_frame, from_=30, to=100, variable=self.quality_var, orient=tk.HORIZONTAL, command=on_quality_drag)
        q_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))
        
        q_entry = ttk.Entry(rq_frame, textvariable=self.quality_var, width=6)
        q_entry.pack(side=tk.LEFT)
        ttk.Label(rq_frame, text="%").pack(side=tk.LEFT)
        
        # 5. Output Directory
        ttk.Label(main_frame, text="5. 选择输出文件夹 (同名防覆盖):").pack(anchor=tk.W, pady=(0,5))
        out_frame = ttk.Frame(main_frame)
        out_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.out_var = tk.StringVar()
        ttk.Entry(out_frame, textvariable=self.out_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(out_frame, text="浏览...", command=self.select_output).pack(side=tk.RIGHT)
        
        # 6. Execution & Progress
        self.start_btn = ttk.Button(main_frame, text="开始处理", command=self.start_processing)
        # Apply standard accent style if available, otherwise fallback
        try:
            self.start_btn.config(style="Accent.TButton")
        except tk.TclError:
            pass
        self.start_btn.pack(fill=tk.X, pady=(0, 10))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        
        self.status_var = tk.StringVar(value="准备就绪")
        ttk.Label(main_frame, textvariable=self.status_var).pack(anchor=tk.W)
        
    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择 PPT 文件",
            filetypes=[("PowerPoint Files", "*.ppt *.pptx")]
        )
        for f in files:
            if f not in self.files:
                self.files.append(f)
                self.file_listbox.insert(tk.END, os.path.basename(f))
                
    def clear_files(self):
        self.files = []
        self.file_listbox.delete(0, tk.END)
        
    def select_watermark(self):
        path = filedialog.askopenfilename(
            title="选择水印图片",
            filetypes=[("Image Files", "*.png *.jpg *.jpeg")]
        )
        if path:
            self.watermark_path = path
            self.wm_var.set(path)
            
    def select_output(self):
        path = filedialog.askdirectory(title="选择输出文件夹")
        if path:
            self.output_dir = path
            self.out_var.set(path)
            
    def start_processing(self):
        if not self.files:
            messagebox.showwarning("警告", "请先添加要处理的 PPT 文件！")
            return
        if not self.output_dir:
            messagebox.showwarning("警告", "请选择输出文件夹！")
            return
            
        try:
            # Parse floats and allow up to 2 decimals for scale
            scale_val = round(float(self.scale_var.get()), 2)
            # JPG Quality is inherently an integer in Pillow
            quality_val = int(round(float(self.quality_var.get())))
        except ValueError:
            messagebox.showwarning("警告", "水印比例和图片质量必须是有效的数字！")
            return
            
        self.start_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)
        
        # We start the background thread
        self.task_manager.start_processing(
            files=self.files,
            output_dir=self.output_dir,
            watermark_path=self.watermark_path,
            scale_percent=scale_val,
            resolution_mode=self.res_var.get(),
            quality=quality_val
        )
        
    def update_progress(self, message, current_file_idx, total_files):
        # Update safely from thread
        def _update():
            self.status_var.set(message)
            if total_files > 0:
                progress = (current_file_idx / total_files) * 100
                self.progress_var.set(progress)
        self.after(0, _update)
        
    def on_finish(self):
        def _finish():
            self.start_btn.config(state=tk.NORMAL)
            self.progress_var.set(100)
            messagebox.showinfo("完成", "所有文件处理完毕！\n体积压缩且防编辑的PPT已生成。")
        self.after(0, _finish)

if __name__ == "__main__":
    app = AppGUI()
    app.mainloop()
