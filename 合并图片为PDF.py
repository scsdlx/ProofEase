# GenPdfFromPIC_UI.py
# 根据选择的图片生成PDF (带UI界面)

import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image

class ImageToPdfApp(tk.Tk):
    def __init__(self):
        super().__init__()

        # --- 实例变量 ---
        self.CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".image_to_pdf_config.json")
        self.title("图片批量转PDF工具 v1.0")
        self.geometry("600x600")

        # --- 参数变量 ---
        self.save_option_var = tk.StringVar(value="same")
        self.output_dir_var = tk.StringVar()
        
        # 新增：图片尺寸选项
        self.resize_option_var = tk.StringVar(value="original")
        self.image_width_var = tk.StringVar()
        self.image_height_var = tk.StringVar()

        # --- 状态与日志变量 ---
        self.log_messages = []

        self._create_widgets()
        self._load_config()
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.LabelFrame(main_frame, text="1. 选择图片文件", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        ttk.Button(file_frame, text="浏览文件", command=self.browse_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="浏览目录", command=self.browse_directory).pack(side=tk.LEFT, padx=5)

        list_frame = ttk.LabelFrame(main_frame, text="2. 待处理图片列表", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.file_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, height=8)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        list_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=list_scrollbar.set)
        
        list_btn_frame = ttk.Frame(list_frame)
        list_btn_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Button(list_btn_frame, text="删除", command=self.remove_selected).pack(pady=2)
        ttk.Button(list_btn_frame, text="清空", command=self.clear_list).pack(pady=2)

        # --- 新的PDF选项区域 ---
        options_frame = ttk.LabelFrame(main_frame, text="3. PDF 选项", padding="10")
        options_frame.pack(fill=tk.X, pady=5)
        
        ttk.Radiobutton(options_frame, text="使用原始图片尺寸", variable=self.resize_option_var, 
                        value="original", command=self.toggle_resize_entries).pack(anchor=tk.W)

        resize_frame = ttk.Frame(options_frame)
        resize_frame.pack(fill=tk.X, anchor=tk.W, pady=(5,0))
        ttk.Radiobutton(resize_frame, text="统一调整为指定尺寸 (像素):", variable=self.resize_option_var, 
                        value="specific", command=self.toggle_resize_entries).pack(side=tk.LEFT)

        self.width_entry = ttk.Entry(resize_frame, textvariable=self.image_width_var, width=8, state="disabled")
        self.width_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(resize_frame, text="x").pack(side=tk.LEFT)
        self.height_entry = ttk.Entry(resize_frame, textvariable=self.image_height_var, width=8, state="disabled")
        self.height_entry.pack(side=tk.LEFT, padx=5)

        output_frame = ttk.LabelFrame(main_frame, text="4. 输出选项", padding="10")
        output_frame.pack(fill=tk.X, pady=5)
        output_frame.columnconfigure(1, weight=1)
        ttk.Radiobutton(output_frame, text="保存到图片所在的原目录", variable=self.save_option_var, value="same", command=self.toggle_output_path).grid(row=0, column=0, columnspan=3, sticky=tk.W)
        ttk.Radiobutton(output_frame, text="保存到指定目录", variable=self.save_option_var, value="specific", command=self.toggle_output_path).grid(row=1, column=0, sticky=tk.W)
        self.output_path_entry = ttk.Entry(output_frame, textvariable=self.output_dir_var, state="disabled")
        self.output_path_entry.grid(row=1, column=1, sticky=tk.EW, padx=5)
        self.btn_browse_output = ttk.Button(output_frame, text="浏览...", command=self.browse_output_dir, state="disabled")
        self.btn_browse_output.grid(row=1, column=2, sticky=tk.E)

        action_frame = ttk.Frame(main_frame, padding="5")
        action_frame.pack(fill=tk.X, pady=5)
        action_frame.columnconfigure(0, weight=1); action_frame.columnconfigure(1, weight=1)
        self.btn_process = ttk.Button(action_frame, text="开始生成PDF", command=self.start_processing)
        self.btn_process.grid(row=0, column=0, sticky=tk.EW, padx=5, ipady=5)
        self.btn_log = ttk.Button(action_frame, text="查看处理日志", command=self._show_log_window)
        self.btn_log.grid(row=0, column=1, sticky=tk.EW, padx=5, ipady=5)
        
        self.status_var = tk.StringVar(value="欢迎使用！请先添加图片。")
        status_label = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=5)
        status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def start_processing(self):
        self.log_messages = []
        self._log("--- 开始新一轮PDF生成任务 ---")
        
        all_files = self.file_listbox.get(0, tk.END)
        if not all_files:
            messagebox.showerror("错误", "待处理列表为空，请先添加图片！")
            self._log("错误: 待处理列表为空。")
            return
            
        # 验证输出目录
        save_option, specific_output_dir = self.save_option_var.get(), self.output_dir_var.get()
        if save_option == "specific" and not os.path.isdir(specific_output_dir):
            messagebox.showerror("错误", "请选择一个有效的指定输出目录！")
            self._log("错误: 指定输出目录无效。")
            return

        # 验证尺寸参数
        resize_option = self.resize_option_var.get()
        target_size = None
        if resize_option == 'specific':
            try:
                target_w = int(self.image_width_var.get())
                target_h = int(self.image_height_var.get())
                if target_w <= 0 or target_h <= 0: raise ValueError
                target_size = (target_w, target_h)
                self._log(f"将统一调整图片尺寸为: {target_w}x{target_h}")
            except (ValueError, TypeError):
                messagebox.showerror("错误", "指定的宽度和高度必须是有效的正整数！")
                self._log("错误: 无效的尺寸参数。")
                return

        # 将文件按父目录分组
        dirs_to_files = {}
        for f_path in all_files:
            dir_path = os.path.dirname(f_path)
            if dir_path not in dirs_to_files:
                dirs_to_files[dir_path] = []
            dirs_to_files[dir_path].append(f_path)
        
        self.btn_process.config(state="disabled")
        total_pdfs = len(dirs_to_files)
        success_count = 0

        for i, (dir_path, image_paths) in enumerate(dirs_to_files.items()):
            dir_name = os.path.basename(dir_path)
            self._log(f"\n({i+1}/{total_pdfs}) 正在处理目录: {dir_name}")
            
            # 按文件名排序，确保图片顺序
            image_paths.sort()
            
            # 定义输出PDF的路径
            if save_option == "specific":
                output_dir = specific_output_dir
            else: # save_option == "same"
                output_dir = dir_path
            
            os.makedirs(output_dir, exist_ok=True)
            pdf_output_path = os.path.join(output_dir, f"{dir_name}.pdf")

            self._log(f"  -> 找到了 {len(image_paths)} 张图片，准备合并...")
            
            images_to_convert = []
            try:
                for path in image_paths:
                    img = Image.open(path)
                    # PDF不支持RGBA（带透明度）模式，需要转换为RGB
                    if img.mode == 'RGBA':
                        img = img.convert('RGB')
                    
                    # 如果需要，调整图片大小
                    if target_size:
                        img = img.resize(target_size, Image.Resampling.LANCZOS)

                    images_to_convert.append(img)
                
                if not images_to_convert:
                    self._log(f"  -> 警告: 目录 '{dir_name}' 中没有可处理的图片，已跳过。")
                    continue

                first_image = images_to_convert[0]
                
                # 使用第一张图片来保存，并附加其余图片
                first_image.save(
                    pdf_output_path, 
                    "PDF", 
                    resolution=100.0, 
                    save_all=True, 
                    append_images=images_to_convert[1:]
                )
                self._log(f"  -> 成功！PDF已保存至: {pdf_output_path}")
                success_count += 1

            except Exception as e:
                self._log(f"  -> 创建PDF时发生错误: {e}")

        self._log(f"\n--- 处理完成！成功生成 {success_count}/{total_pdfs} 个PDF文件。 ---")
        self.btn_process.config(state="normal")
        messagebox.showinfo("完成", f"所有任务已处理完毕。\n\n成功生成: {success_count} 个PDF\n处理目录数: {total_pdfs}\n\n详细信息请点击“查看处理日志”。")

    def toggle_resize_entries(self):
        state = "normal" if self.resize_option_var.get() == "specific" else "disabled"
        self.width_entry.config(state=state)
        self.height_entry.config(state=state)

    def _update_default_size(self):
        """获取列表第一张图片的尺寸并填充到输入框"""
        all_files = self.file_listbox.get(0, tk.END)
        if not all_files:
            self.image_width_var.set("")
            self.image_height_var.set("")
            return

        first_file = all_files[0]
        try:
            with Image.open(first_file) as img:
                width, height = img.size
                self.image_width_var.set(str(width))
                self.image_height_var.set(str(height))
                self.update_status(f"已使用第一张图片尺寸 {width}x{height} 作为默认值。")
        except Exception as e:
            self.image_width_var.set("")
            self.image_height_var.set("")
            self.update_status(f"无法读取第一张图片尺寸: {e}")
    
    def browse_files(self):
        files = filedialog.askopenfilenames(title="选择图片文件", filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp *.tiff"), ("All files", "*.*")])
        if files:
            for f in files:
                if f not in self.file_listbox.get(0, tk.END):
                    self.file_listbox.insert(tk.END, f)
            self.update_status(f"添加了 {len(files)} 个文件。")
            self._update_default_size()

    def browse_directory(self):
        directory = filedialog.askdirectory(title="选择图片所在目录")
        if not directory: return
        
        count = 0
        allowed_extensions = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".gif"}
        
        # 遍历目录及其子目录
        for root, _, filenames in os.walk(directory):
            for filename in filenames:
                if os.path.splitext(filename)[1].lower() in allowed_extensions:
                    full_path = os.path.join(root, filename)
                    if full_path not in self.file_listbox.get(0, tk.END):
                        self.file_listbox.insert(tk.END, full_path)
                        count += 1
        
        self.update_status(f"从目录及其子目录中添加了 {count} 个文件。")
        self._update_default_size()

    def remove_selected(self):
        for i in sorted(self.file_listbox.curselection(), reverse=True):
            self.file_listbox.delete(i)
        self.update_status("已删除选中项。")
        self._update_default_size()

    def clear_list(self):
        self.file_listbox.delete(0, tk.END)
        self.update_status("列表已清空。")
        self._update_default_size()
    
    def _log(self, message, to_status=True):
        self.log_messages.append(message)
        if to_status: self.update_status(message)
        print(message) 

    def _show_log_window(self):
        log_window = tk.Toplevel(self); log_window.title("处理日志"); log_window.geometry("700x500")
        text_frame = ttk.Frame(log_window); text_frame.pack(expand=True, fill=tk.BOTH, padx=5, pady=5)
        log_text = tk.Text(text_frame, wrap=tk.WORD, state=tk.DISABLED, height=10, width=80)
        log_text.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y); log_text.config(yscrollcommand=scrollbar.set)
        log_text.config(state=tk.NORMAL); log_text.delete(1.0, tk.END); log_text.insert(tk.END, "\n".join(self.log_messages)); log_text.config(state=tk.DISABLED); log_text.see(tk.END)

    def _save_config(self):
        config = {
            "version": "1.0",
            "save_option": self.save_option_var.get(),
            "output_dir": self.output_dir_var.get(),
            "resize_option": self.resize_option_var.get(),
            "image_width": self.image_width_var.get(),
            "image_height": self.image_height_var.get(),
        }
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            self._log(f"无法保存配置: {e}", to_status=False)

    def _load_config(self):
        if not os.path.exists(self.CONFIG_FILE): return
        try:
            with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            self.save_option_var.set(config.get("save_option", "same"))
            self.output_dir_var.set(config.get("output_dir", ""))
            self.resize_option_var.set(config.get("resize_option", "original"))
            self.image_width_var.set(config.get("image_width", ""))
            self.image_height_var.set(config.get("image_height", ""))
            
            self.toggle_output_path()
            self.toggle_resize_entries()
            self.update_status("已加载上次的配置。")
        except Exception as e:
            self._log(f"无法加载配置: {e}", to_status=True)

    def toggle_output_path(self):
        state = "normal" if self.save_option_var.get() == "specific" else "disabled"
        self.output_path_entry.config(state=state)
        self.btn_browse_output.config(state=state)

    def browse_output_dir(self):
        directory = filedialog.askdirectory(title="选择保存目录")
        if directory:
            self.output_dir_var.set(directory)

    def update_status(self, message):
        self.status_var.set(message)
        self.update_idletasks()

    def _on_closing(self):
        self._save_config()
        self.destroy()

if __name__ == '__main__':
    app = ImageToPdfApp()
    app.mainloop()