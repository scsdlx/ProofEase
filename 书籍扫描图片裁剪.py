# 文件：书籍扫描图片裁剪_new.py

# 扫描图片处理：将图片中的内容进行智能裁剪，并保存
import cv2
import numpy as np
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import json
from PIL import Image, ImageTk
import tempfile
import shutil

class BookCropperApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # --- 实例变量 ---
        # V7.4: 增加主界面设置按钮，主界面参数只读，增加独立的参数配置加载/保存功能
        self.CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".book_cropper_config_v7.4.json")
        self.title("书籍页面智能裁剪工具 v7.4")
        self.geometry("650x800")
        self.resizable(False, False)

        # --- 参数变量 (UI在主窗口和预览窗口中都会存在) ---
        self.crop_width_var = tk.StringVar(value="1780")
        self.crop_height_var = tk.StringVar(value="2550")
        self.top_offset_var = tk.StringVar(value="50")
        self.left_margin_var = tk.StringVar(value="160")
        
        self.save_option_var = tk.StringVar(value="same")
        self.output_dir_var = tk.StringVar()
        self.debug_mode_var = tk.BooleanVar(value=True)

        # --- 内容检测参数 ---
        self.NUM_BG_COLORS = 5
        self.bg_colors = [None] * self.NUM_BG_COLORS
        self.bg_tolerance_vars = [tk.IntVar(value=25) for _ in range(self.NUM_BG_COLORS)]
        self.bg_enabled_vars = [tk.BooleanVar(value=True) for _ in range(self.NUM_BG_COLORS)]
        self.active_swatch_index = -1

        self.expansion_var = tk.StringVar(value="6")
        self.exclude_aspect_ratio_var = tk.BooleanVar(value=True)
        self.aspect_ratio_threshold_var = tk.StringVar(value="6")
        self.clear_edges_var = tk.BooleanVar(value=True)
        self.edge_width_var = tk.StringVar(value="30")
        self.min_area_ratio_var = tk.StringVar(value="1.0")
        
        # --- 状态与日志变量 ---
        self.preview_window = None
        self.log_messages = []
        
        # --- UI 控件与预览状态引用 ---
        self.main_swatch_buttons = []
        self.preview_widgets = {}
        self.zoom_level = 1.0
        self.zoom_var = tk.DoubleVar(value=1.0)
        self.zoom_label_var = tk.StringVar()
        self.image_on_canvas = None
        self.current_preview_cv_image = None
        self.original_preview_image = None
        self.preview_current_index = 0 # 预览窗口内使用的索引

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

        list_frame = ttk.LabelFrame(main_frame, text="2. 待处理图片列表 (双击或选中后点击'设置')", padding="10")
        list_frame.pack(fill=tk.X, pady=5)
        self.file_listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, height=8)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.file_listbox.bind("<Double-Button-1>", self._on_listbox_select)
        self.file_listbox.bind("<<ListboxSelect>>", self._on_listbox_selection_change) # 新增绑定
        list_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=list_scrollbar.set)
        list_btn_frame = ttk.Frame(list_frame)
        list_btn_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Button(list_btn_frame, text="删除", command=self.remove_selected).pack(pady=2)
        ttk.Button(list_btn_frame, text="清空", command=self.clear_list).pack(pady=2)
        # --- 新增“设置”按钮 ---
        self.btn_settings = ttk.Button(list_btn_frame, text="设置", command=self._open_preview_window, state=tk.DISABLED)
        self.btn_settings.pack(pady=2)


        params_frame = ttk.LabelFrame(main_frame, text="3. 当前参数预览 (请在设置窗口中修改)", padding="10")
        params_frame.pack(fill=tk.X, pady=5)
        
        bg_frame = ttk.LabelFrame(params_frame, text="内容检测: 背景色", padding=5)
        bg_frame.pack(fill=tk.X, pady=2)
        bg_slots_container = ttk.Frame(bg_frame)
        bg_slots_container.pack()
        for i in range(self.NUM_BG_COLORS):
            btn = tk.Button(bg_slots_container, text="空", width=8, relief=tk.RAISED, state="disabled")
            btn.grid(row=0, column=i, padx=5, pady=2)
            self.main_swatch_buttons.append(btn)

        other_params_frame = ttk.LabelFrame(params_frame, text="内容检测: 其他参数", padding=10)
        other_params_frame.pack(fill=tk.X, pady=(10, 2))
        other_params_frame.columnconfigure(1, weight=1); other_params_frame.columnconfigure(3, weight=1)
        ttk.Checkbutton(other_params_frame, text="清除边缘(px):", variable=self.clear_edges_var, state="disabled").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(other_params_frame, textvariable=self.edge_width_var, width=8, state="readonly").grid(row=0, column=1, sticky=tk.EW)
        ttk.Label(other_params_frame, text="选区扩展(px):").grid(row=0, column=2, sticky=tk.E)
        ttk.Entry(other_params_frame, textvariable=self.expansion_var, width=8, state="readonly").grid(row=0, column=3, sticky=tk.EW, padx=5)
        ttk.Label(other_params_frame, text="最小面积(‱):").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(other_params_frame, textvariable=self.min_area_ratio_var, width=8, state="readonly").grid(row=1, column=1, sticky=tk.EW)
        ttk.Checkbutton(other_params_frame, text="排除宽高比>", variable=self.exclude_aspect_ratio_var, state="disabled").grid(row=1, column=2, sticky=tk.E)
        ttk.Entry(other_params_frame, textvariable=self.aspect_ratio_threshold_var, width=8, state="readonly").grid(row=1, column=3, sticky=tk.EW, padx=5)

        crop_size_frame = ttk.LabelFrame(params_frame, text="最终裁剪参数 (单位: 像素)", padding="10")
        crop_size_frame.pack(fill=tk.X, pady=2)
        crop_size_frame.columnconfigure(1, weight=1); crop_size_frame.columnconfigure(3, weight=1)
        ttk.Label(crop_size_frame, text="宽度:").grid(row=0, column=0, sticky=tk.W); ttk.Entry(crop_size_frame, textvariable=self.crop_width_var, state="readonly").grid(row=0, column=1, sticky=tk.EW, padx=(0,10))
        ttk.Label(crop_size_frame, text="高度:").grid(row=0, column=2, sticky=tk.W); ttk.Entry(crop_size_frame, textvariable=self.crop_height_var, state="readonly").grid(row=0, column=3, sticky=tk.EW)
        ttk.Label(crop_size_frame, text="顶部偏移:").grid(row=1, column=0, sticky=tk.W); ttk.Entry(crop_size_frame, textvariable=self.top_offset_var, state="readonly").grid(row=1, column=1, sticky=tk.EW, padx=(0,10))
        ttk.Label(crop_size_frame, text="左边距:").grid(row=1, column=2, sticky=tk.W); ttk.Entry(crop_size_frame, textvariable=self.left_margin_var, state="readonly").grid(row=1, column=3, sticky=tk.EW)
        
        output_frame = ttk.LabelFrame(main_frame, text="4. 输出与调试", padding="10")
        output_frame.pack(fill=tk.X, pady=5)
        output_frame.columnconfigure(1, weight=1)
        ttk.Radiobutton(output_frame, text="保存到原目录", variable=self.save_option_var, value="same", command=self.toggle_output_path, state="disabled").grid(row=0, column=0, columnspan=3, sticky=tk.W)
        ttk.Radiobutton(output_frame, text="保存到指定目录", variable=self.save_option_var, value="specific", command=self.toggle_output_path, state="disabled").grid(row=1, column=0, sticky=tk.W)
        self.output_path_entry = ttk.Entry(output_frame, textvariable=self.output_dir_var, state="disabled")
        self.output_path_entry.grid(row=1, column=1, sticky=tk.EW, padx=5)
        self.btn_browse_output = ttk.Button(output_frame, text="浏览...", command=self.browse_output_dir, state="disabled")
        self.btn_browse_output.grid(row=1, column=2, sticky=tk.E)
        ttk.Checkbutton(output_frame, text="生成调试图片", variable=self.debug_mode_var, state="disabled").grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        action_frame = ttk.Frame(main_frame, padding="5")
        action_frame.pack(fill=tk.X, pady=5)
        action_frame.columnconfigure(0, weight=1); action_frame.columnconfigure(1, weight=1); action_frame.columnconfigure(2, weight=1)
        self.btn_process = ttk.Button(action_frame, text="开始处理", command=self.start_processing)
        self.btn_process.grid(row=0, column=0, sticky=tk.EW, padx=5, ipady=5)
        self.btn_show_steps = ttk.Button(action_frame, text="显示处理过程", command=self._show_processing_steps)
        self.btn_show_steps.grid(row=0, column=1, sticky=tk.EW, padx=5, ipady=5)
        self.btn_log = ttk.Button(action_frame, text="查看处理日志", command=self._show_log_window)
        self.btn_log.grid(row=0, column=2, sticky=tk.EW, padx=5, ipady=5)
        
        self.status_var = tk.StringVar(value="欢迎使用！请先添加图片。")
        status_label = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=5)
        status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def start_processing(self):
        # ... (此函数保持不变)
        self.log_messages = []
        self._log("--- 开始新一轮处理任务 ---")
        files_to_process = self.file_listbox.get(0, tk.END)
        if not files_to_process:
            messagebox.showerror("错误", "待处理列表为空，请先添加图片！"); self._log("错误: 待处理列表为空。"); return
        active_colors = []
        for i in range(self.NUM_BG_COLORS):
            if self.bg_colors[i] is not None and self.bg_enabled_vars[i].get():
                active_colors.append({"color": self.bg_colors[i], "tolerance": self.bg_tolerance_vars[i].get()})
        if not active_colors:
            messagebox.showerror("错误", "请至少设置并启用一个有效的背景色！"); self._log("错误: 未设置或启用有效背景色。"); return
        try:
            params = {
                "crop_w": int(self.crop_width_var.get()), "crop_h": int(self.crop_height_var.get()),
                "offset_y": int(self.top_offset_var.get()), "left_margin": int(self.left_margin_var.get()),
                "expansion": int(self.expansion_var.get()), "aspect_ratio_limit": float(self.aspect_ratio_threshold_var.get()),
                "edge_width": int(self.edge_width_var.get()), "min_area_ratio": float(self.min_area_ratio_var.get())
            }
            if any(v < 0 for k, v in params.items() if k != "aspect_ratio_limit") or params["aspect_ratio_limit"] <= 0: raise ValueError
        except ValueError:
            messagebox.showerror("错误", "所有数值参数必须是有效的正数（部分可为0）！"); self._log("错误: 参数无效。"); return
        save_option, specific_output_dir = self.save_option_var.get(), self.output_dir_var.get()
        if save_option == "specific" and not os.path.isdir(specific_output_dir):
            messagebox.showerror("错误", "请选择一个有效的指定输出目录！"); self._log("错误: 指定输出目录无效。"); return
        self.btn_process.config(state="disabled")
        total_files, success_count = len(files_to_process), 0
        for i, image_path in enumerate(files_to_process):
            basename = os.path.basename(image_path)
            self._log(f"({i+1}/{total_files}) 开始处理: {basename}")
            try:
                source_dir, filename = os.path.split(image_path); name, ext = os.path.splitext(filename)
                final_output_dir = specific_output_dir if save_option == "specific" else source_dir
                os.makedirs(final_output_dir, exist_ok=True)
                final_cropped_path = os.path.join(final_output_dir, f"{name}_cropped{ext}")
                debug_base_path = None
                if self.debug_mode_var.get():
                    debug_dir = os.path.join(final_output_dir, "debug_output"); os.makedirs(debug_dir, exist_ok=True)
                    debug_base_path = os.path.join(debug_dir, name); self._log(f"  > 调试模式开启", to_status=False)
                with open(image_path, 'rb') as f: image_data = np.frombuffer(f.read(), np.uint8)
                original_image = cv2.imdecode(image_data, cv2.IMREAD_COLOR)
                if original_image is None: self._log(f"  > 警告: 无法解码图片 {basename}"); continue
                img_h, img_w = original_image.shape[:2]; self._log(f"  > 图片加载成功: {img_w}x{img_h}", to_status=False)
                if debug_base_path: self._save_image_robust(f"{debug_base_path}_01_original.png", original_image)
                working_image, angle = self._deskew_with_projection_profile(original_image)
                if debug_base_path: self._save_image_robust(f"{debug_base_path}_02_rotated.png", working_image)
                img_h, img_w = working_image.shape[:2]
                if self.clear_edges_var.get() and params["edge_width"] > 0:
                    ew = min(params["edge_width"], img_h // 2, img_w // 2); color_bgr = active_colors[0]['color'].tolist()
                    cv2.rectangle(working_image, (0, 0), (img_w - 1, ew - 1), color_bgr, -1); cv2.rectangle(working_image, (0, img_h - ew), (img_w - 1, img_h - 1), color_bgr, -1); cv2.rectangle(working_image, (0, 0), (ew - 1, img_h - 1), color_bgr, -1); cv2.rectangle(working_image, (img_w - ew, 0), (img_w - 1, img_h - 1), color_bgr, -1)
                    if debug_base_path: self._save_image_robust(f"{debug_base_path}_03_edges_cleared.png", working_image)
                total_background_mask = np.zeros((img_h, img_w), dtype=np.uint8)
                for ac in active_colors:
                    lower = np.clip(ac["color"].astype(np.int16) - ac["tolerance"], 0, 255).astype(np.uint8)
                    upper = np.clip(ac["color"].astype(np.int16) + ac["tolerance"], 0, 255).astype(np.uint8)
                    total_background_mask = cv2.bitwise_or(total_background_mask, cv2.inRange(working_image, lower, upper))
                content_mask = cv2.bitwise_not(total_background_mask)
                if debug_base_path: self._save_image_robust(f"{debug_base_path}_04_content_mask.png", content_mask)
                if params["expansion"] > 0:
                    kernel = np.ones((params["expansion"], params["expansion"]), np.uint8)
                    content_mask = cv2.dilate(content_mask, kernel, iterations=1)
                    if debug_base_path: self._save_image_robust(f"{debug_base_path}_05_dilated_mask.png", content_mask)
                contours, _ = cv2.findContours(content_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                min_area = (img_h * img_w) * (params["min_area_ratio"] / 10000.0)
                significant_contours = []
                for c in contours:
                    if cv2.contourArea(c) < min_area: continue
                    if self.exclude_aspect_ratio_var.get():
                        x, y, w, h = cv2.boundingRect(c)
                        if w == 0 or h == 0 or max(w/h, h/w) > params["aspect_ratio_limit"]: continue
                    significant_contours.append(c)
                if debug_base_path:
                    dbg_img_contours = working_image.copy()
                    cv2.drawContours(dbg_img_contours, significant_contours, -1, (0, 0, 255), 3)
                    self._save_image_robust(f"{debug_base_path}_06_filtered_contours.png", dbg_img_contours)
                if not significant_contours: self._log(f"  > 警告: 未找到有效内容。"); continue
                all_points = np.vstack(significant_contours)
                content_box_x_min, _, content_width, _ = cv2.boundingRect(all_points)
                if content_width >= params["crop_w"] * 0.9:
                    self._log(f"  > 内容宽度({content_width}px)较大，采用居中策略。", to_status=False)
                    content_center_x = content_box_x_min + content_width / 2
                    x1 = content_center_x - (params["crop_w"] / 2)
                else:
                    self._log(f"  > 采用左边距策略 (边距: {params['left_margin']}px)。", to_status=False)
                    x1 = all_points[:, :, 0].min() - params["left_margin"]
                x1 = max(0, x1); y1 = max(0, params["offset_y"]); x2 = min(img_w, x1 + params["crop_w"]); y2 = min(img_h, y1 + params["crop_h"])
                if debug_base_path:
                    dbg_img_crop_box = working_image.copy()
                    cv2.rectangle(dbg_img_crop_box, (int(x1), int(y1)), (int(x2), int(y2)), (0, 255, 0), 5)
                    self._save_image_robust(f"{debug_base_path}_07_crop_area.png", dbg_img_crop_box)
                cropped_image = working_image[int(y1):int(y2), int(x1):int(x2)]
                if debug_base_path:
                    self._save_image_robust(f"{debug_base_path}_08_final_cropped.png", cropped_image)
                self._save_image_robust(final_cropped_path, cropped_image, ext)
                self._log(f"  > 保存成功: {os.path.basename(final_cropped_path)}", to_status=False)
                success_count += 1
            except Exception as e:
                import traceback; traceback.print_exc()
                self._log(f"  > 错误: 处理文件 {basename} 时发生异常: {e}")
        self._log(f"--- 处理完成！成功 {success_count}/{total_files}。 ---")
        self.btn_process.config(state="normal")
        messagebox.showinfo("完成", f"所有任务已处理完毕。\n\n成功: {success_count}\n失败: {total_files - success_count}\n\n详细信息请点击“查看处理日志”。")

    def _deskew_with_projection_profile(self, img):
        # ... (此函数保持不变)
        self._log("  > 使用投影剖面法进行倾斜校正...", to_status=False)
        img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, img_binary = cv2.threshold(img_gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        h, w = img_binary.shape[:2]
        angle_range = 5; angle_step = 0.1
        angles = np.arange(-angle_range, angle_range + angle_step, angle_step)
        max_score = -1.0; best_angle = 0.0
        for angle in angles:
            center = (w // 2, h // 2)
            M = cv2.getRotationMatrix2D(center, angle, 1.0)
            rotated_binary = cv2.warpAffine(img_binary, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_CONSTANT, borderValue=0)
            horizontal_projection = np.sum(rotated_binary, axis=1)
            score = np.var(horizontal_projection)
            if score > max_score:
                max_score = score
                best_angle = angle
        self._log(f"  > 投影法检测到最佳倾斜角: {best_angle:.2f}°", to_status=False)
        if abs(best_angle) < 0.1:
             self._log(f"  > 倾斜角度 {best_angle:.2f}° 过小，跳过旋转。", to_status=False)
             return img, 0.0
        if abs(best_angle) >= angle_range:
             self._log(f"  > 警告: 倾斜角度达到搜索边界 ({best_angle:.2f}°), 可能校正不准确。", to_status=False)
        self._log(f"  > 进行校正...", to_status=False)
        center = (img.shape[1] // 2, img.shape[0] // 2)
        M_final = cv2.getRotationMatrix2D(center, best_angle, 1.0)
        fill_color = self.bg_colors[0].tolist() if self.bg_colors[0] is not None and len(self.bg_colors) > 0 else [255, 255, 255]
        rotated_img = cv2.warpAffine(img, M_final, (img.shape[1], img.shape[0]), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_CONSTANT, borderValue=fill_color)
        return rotated_img, best_angle

    # --- 以下为UI、配置、预览窗口相关函数 (已重构) ---

    def _update_all_slot_uis(self, index):
        # ... (此函数保持不变)
        color = self.bg_colors[index]
        main_btn = self.main_swatch_buttons[index]
        hex_color, fg_color = "SystemButtonFace", "SystemButtonText"
        btn_text = "空"

        if color is not None:
            tolerance = self.bg_tolerance_vars[index].get()
            hex_color = f'#{color[2]:02x}{color[1]:02x}{color[0]:02x}'
            fg_color = self._get_contrasting_text_color(color)
            btn_text = f"T:{tolerance}"
        
        main_btn.config(bg=hex_color, text=btn_text, fg=fg_color)

        if self.preview_window and self.preview_window.winfo_exists():
            preview_btn = self.preview_widgets['swatches'][index]
            slot_frame = self.preview_widgets['frames'][index]
            preview_btn.config(bg=hex_color, text=btn_text, fg=fg_color)
            is_active = (index == self.active_swatch_index)
            slot_frame.config(relief=tk.RIDGE if is_active else tk.FLAT, borderwidth=3 if is_active else 1)

    def _get_contrasting_text_color(self, bgr_color):
        # ... (此函数保持不变)
        b, g, r = bgr_color
        brightness = (int(r) * 299 + int(g) * 587 + int(b) * 114) / 1000
        return 'white' if brightness < 128 else 'black'

    def _log(self, message, to_status=True):
        # ... (此函数保持不变)
        self.log_messages.append(message)
        if to_status: self.update_status(message)
        print(message) 

    def _show_log_window(self):
        # ... (此函数保持不变)
        log_window = tk.Toplevel(self); log_window.title("处理日志"); log_window.geometry("700x500")
        text_frame = ttk.Frame(log_window); text_frame.pack(expand=True, fill=tk.BOTH, padx=5, pady=5)
        log_text = tk.Text(text_frame, wrap=tk.WORD, state=tk.DISABLED, height=10, width=80)
        log_text.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y); log_text.config(yscrollcommand=scrollbar.set)
        log_text.config(state=tk.NORMAL); log_text.delete(1.0, tk.END); log_text.insert(tk.END, "\n".join(self.log_messages)); log_text.config(state=tk.DISABLED); log_text.see(tk.END)
    
    # --- 新增参数配置的获取和应用函数 ---
    def _get_parameters_as_dict(self):
        """将所有可配置参数打包成一个字典"""
        colors_config = []
        for i in range(self.NUM_BG_COLORS):
            colors_config.append({
                "color": self.bg_colors[i].tolist() if self.bg_colors[i] is not None else None,
                "tolerance": self.bg_tolerance_vars[i].get(),
                "enabled": self.bg_enabled_vars[i].get()
            })
        
        params = {
            "version": "7.4-profile",
            "crop_width": self.crop_width_var.get(), "crop_height": self.crop_height_var.get(),
            "top_offset": self.top_offset_var.get(), "left_margin": self.left_margin_var.get(),
            "save_option": self.save_option_var.get(), "output_dir": self.output_dir_var.get(),
            "debug_mode": self.debug_mode_var.get(), "background_colors_v6_2": colors_config,
            "expansion": self.expansion_var.get(), "exclude_aspect_ratio": self.exclude_aspect_ratio_var.get(),
            "aspect_ratio_threshold": self.aspect_ratio_threshold_var.get(), "clear_edges": self.clear_edges_var.get(),
            "edge_width": self.edge_width_var.get(), "min_area_ratio": self.min_area_ratio_var.get(),
        }
        return params
        
    def _apply_parameters_from_dict(self, config):
        """从字典加载并应用参数"""
        self.crop_width_var.set(config.get("crop_width", "1780")); self.crop_height_var.set(config.get("crop_height", "2550"))
        self.top_offset_var.set(config.get("top_offset", "50")); self.left_margin_var.set(config.get("left_margin", "160"))
        self.save_option_var.set(config.get("save_option", "same")); self.output_dir_var.set(config.get("output_dir", ""))
        self.debug_mode_var.set(config.get("debug_mode", True))
        colors_config = config.get("background_colors_v6_2", [])
        for i in range(self.NUM_BG_COLORS):
            if i < len(colors_config):
                c_conf = colors_config[i]
                self.bg_tolerance_vars[i].set(c_conf.get("tolerance", 25))
                self.bg_enabled_vars[i].set(c_conf.get("enabled", True))
                color_val = c_conf.get("color")
                self.bg_colors[i] = np.array(color_val, dtype=np.uint8) if color_val else None
            else: # 如果配置文件颜色少于槽位，则重置多余的槽位
                self.bg_colors[i] = None
                self.bg_tolerance_vars[i].set(25)
                self.bg_enabled_vars[i].set(True)
            self._update_all_slot_uis(i)
        self.expansion_var.set(config.get("expansion", "6")); self.exclude_aspect_ratio_var.set(config.get("exclude_aspect_ratio", True))
        self.aspect_ratio_threshold_var.set(config.get("aspect_ratio_threshold", "6")); self.clear_edges_var.set(config.get("clear_edges", True))
        self.edge_width_var.set(config.get("edge_width", "30")); self.min_area_ratio_var.set(config.get("min_area_ratio", "1.0"))
        self.toggle_output_path()
        self.update_status("已成功加载参数配置。")

    def _save_config(self):
        config = self._get_parameters_as_dict()
        # 添加非参数配置
        config["version"] = "7.4"
        config["zoom_level"] = self.zoom_level
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump(config, f, indent=4)
        except Exception as e: self._log(f"无法保存配置: {e}", to_status=False)

    def _load_config(self):
        if not os.path.exists(self.CONFIG_FILE): return
        try:
            with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f: config = json.load(f)
            self._apply_parameters_from_dict(config)
            self.zoom_level = config.get("zoom_level", 1.0)
            self.zoom_var.set(self.zoom_level)
            self.update_status("已加载上次的配置。")
        except Exception as e: self._log(f"无法加载配置: {e}", to_status=True)

    def _open_preview_window(self):
        # MODAL IMPLEMENTATION
        selection = self.file_listbox.curselection()
        if not selection:
            messagebox.showinfo("提示", "请先在列表中选择一张图片。")
            return
            
        self.preview_current_index = selection[0]

        self.preview_window = tk.Toplevel(self)
        self.preview_window.title("预览与参数设置 [Ctrl+滚轮:缩放, 滚轮:垂直滚动, Shift+滚轮:水平滚动]")
        self.preview_window.geometry("1200x800")
        self.preview_window.protocol("WM_DELETE_WINDOW", self._on_preview_close)

        # Create UI
        main_pane = ttk.PanedWindow(self.preview_window, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True)
        view_area_frame = ttk.Frame(main_pane)
        self._create_canvas_area(view_area_frame) 
        self._create_zoom_controls(view_area_frame) 
        self._create_nav_controls(view_area_frame)
        main_pane.add(view_area_frame, weight=3)
        controls_container = ttk.Frame(main_pane)
        controls_canvas = tk.Canvas(controls_container)
        controls_scrollbar = ttk.Scrollbar(controls_container, orient="vertical", command=controls_canvas.yview)
        controls_frame = ttk.Frame(controls_canvas, padding=10)
        controls_frame.bind("<Configure>", lambda e: controls_canvas.configure(scrollregion=controls_canvas.bbox("all")))
        controls_canvas.create_window((0, 0), window=controls_frame, anchor="nw")
        controls_canvas.configure(yscrollcommand=controls_scrollbar.set)
        controls_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        controls_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        main_pane.add(controls_container, weight=1)
        self._populate_preview_controls(controls_frame)

        # Load initial image and apply preview
        if self._load_image_for_preview(self.preview_current_index):
            for i in range(self.NUM_BG_COLORS): self._update_all_slot_uis(i)
            self._update_nav_buttons_state()
            self._reset_view()
        
        # Make it modal
        self.preview_window.grab_set()
        self.wait_window(self.preview_window)
        
    def _load_image_for_preview(self, index):
        try:
            selected_path = self.file_listbox.get(index)
            with open(selected_path, 'rb') as f:
                image_data = np.frombuffer(f.read(), np.uint8)
            self.original_preview_image = cv2.imdecode(image_data, cv2.IMREAD_COLOR)
            if self.original_preview_image is None:
                raise IOError("无法解码图像")
            return True
        except Exception as e:
            messagebox.showerror("错误", f"无法打开图片: {e}")
            self._on_preview_close()
            return False

    def _create_canvas_area(self, parent):
        canvas_frame = ttk.Frame(parent)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(canvas_frame, bg="gray50", cursor="crosshair")
        v_scroll = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        h_scroll = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.config(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.canvas.bind("<MouseWheel>", self._on_zoom)
        self.canvas.bind("<Button-4>", self._on_zoom); self.canvas.bind("<Button-5>", self._on_zoom)
        self.canvas.bind("<ButtonPress-2>", self._on_pan_start); self.canvas.bind("<B2-Motion>", self._on_pan_drag)
        self.canvas.bind("<ButtonPress-3>", self._on_pan_start); self.canvas.bind("<B3-Motion>", self._on_pan_drag)
        self.canvas.bind("<ButtonPress-1>", self._on_preview_press)
        self.canvas.bind("<ButtonRelease-1>", self._on_preview_release_and_pick)

    def _create_zoom_controls(self, parent):
        zoom_frame = tk.Frame(parent)
        zoom_frame.place(relx=1.0, rely=0.5, anchor=tk.E, x=-10)

        self.btn_zoom_in = tk.Button(zoom_frame, text="+", command=self._zoom_in, width=2, relief=tk.FLAT)
        self.btn_zoom_in.pack(pady=(0, 2))
        
        zoom_slider = ttk.Scale(zoom_frame, from_=8.0, to=0.1, variable=self.zoom_var, orient=tk.VERTICAL, command=self._on_zoom_scale_change, length=150)
        zoom_slider.pack(pady=2, fill=tk.Y, expand=True)
        
        self.btn_zoom_out = tk.Button(zoom_frame, text="-", command=self._zoom_out, width=2, relief=tk.FLAT)
        self.btn_zoom_out.pack(pady=2)
        
        zoom_label = ttk.Label(zoom_frame, textvariable=self.zoom_label_var, anchor=tk.CENTER)
        zoom_label.pack(pady=(2, 0))

        self._update_zoom_label()

    def _create_nav_controls(self, parent):
        nav_frame = tk.Frame(parent)
        nav_frame.place(relx=0.5, rely=0.95, anchor=tk.S)

        self.btn_prev = tk.Button(nav_frame, text="< 上一幅", command=lambda: self._navigate_preview(-1), relief=tk.FLAT)
        self.btn_prev.pack(side=tk.LEFT, padx=10)
        
        self.btn_next = tk.Button(nav_frame, text="下一幅 >", command=lambda: self._navigate_preview(1), relief=tk.FLAT)
        self.btn_next.pack(side=tk.LEFT, padx=10)

    def _update_nav_buttons_state(self):
        if not (self.preview_window and self.preview_window.winfo_exists()):
            return
            
        list_size = self.file_listbox.size()
        self.btn_prev.config(state=tk.NORMAL if self.preview_current_index > 0 else tk.DISABLED)
        self.btn_next.config(state=tk.NORMAL if self.preview_current_index < list_size - 1 else tk.DISABLED)

    def _navigate_preview(self, direction):
        new_index = self.preview_current_index + direction
        list_size = self.file_listbox.size()
        
        if 0 <= new_index < list_size:
            self.preview_current_index = new_index
            if self._load_image_for_preview(self.preview_current_index):
                self._update_nav_buttons_state()
                self._reset_view()
            
    def _populate_preview_controls(self, parent_frame):
        self.preview_widgets = {'frames': [], 'swatches': []}
        
        # --- 新增：参数配置加载/保存区域 ---
        profile_frame = ttk.LabelFrame(parent_frame, text="参数配置文件", padding=5)
        profile_frame.pack(fill=tk.X, pady=(5, 10))
        profile_frame.columnconfigure(0, weight=1)
        profile_frame.columnconfigure(1, weight=1)
        ttk.Button(profile_frame, text="加载参数配置", command=self._load_parameter_profile).grid(row=0, column=0, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(profile_frame, text="保存参数配置", command=self._save_parameter_profile).grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        
        bg_settings_frame = ttk.LabelFrame(parent_frame, text="背景色设置", padding=5)
        bg_settings_frame.pack(fill=tk.X, pady=5)
        for i in range(self.NUM_BG_COLORS):
            slot_frame = ttk.Frame(bg_settings_frame, padding=(5, 3))
            slot_frame.pack(fill=tk.X, pady=1)
            self.preview_widgets['frames'].append(slot_frame)
            cb = ttk.Checkbutton(slot_frame, text=f"背景 {i+1}", variable=self.bg_enabled_vars[i], command=lambda idx=i: self._update_all_slot_uis(idx))
            cb.pack(side=tk.LEFT, padx=(0, 5))
            btn = tk.Button(slot_frame, text="空", width=8, relief=tk.RAISED, command=lambda idx=i: self._set_active_slot_from_preview(idx))
            btn.pack(side=tk.LEFT, padx=5); self.preview_widgets['swatches'].append(btn)
            
            tolerance_frame = ttk.Frame(slot_frame)
            tolerance_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            entry = ttk.Entry(tolerance_frame, textvariable=self.bg_tolerance_vars[i], width=5)
            entry.pack(side=tk.RIGHT)
            slider = ttk.Scale(tolerance_frame, from_=0, to=100, variable=self.bg_tolerance_vars[i], command=lambda v, idx=i: self._on_tolerance_change(idx))
            slider.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=5)
            entry.bind("<Return>", lambda e, idx=i: self._on_tolerance_change(idx))

        crop_size_frame = ttk.LabelFrame(parent_frame, text="最终裁剪参数 (单位: 像素)", padding="10")
        crop_size_frame.pack(fill=tk.X, pady=5); crop_size_frame.columnconfigure(1, weight=1); crop_size_frame.columnconfigure(3, weight=1)
        ttk.Label(crop_size_frame, text="宽度:").grid(row=0, column=0, sticky=tk.W); ttk.Entry(crop_size_frame, textvariable=self.crop_width_var).grid(row=0, column=1, sticky=tk.EW, padx=(0,10))
        ttk.Label(crop_size_frame, text="高度:").grid(row=0, column=2, sticky=tk.W); ttk.Entry(crop_size_frame, textvariable=self.crop_height_var).grid(row=0, column=3, sticky=tk.EW)
        ttk.Label(crop_size_frame, text="顶部偏移:").grid(row=1, column=0, sticky=tk.W); ttk.Entry(crop_size_frame, textvariable=self.top_offset_var).grid(row=1, column=1, sticky=tk.EW, padx=(0,10))
        ttk.Label(crop_size_frame, text="左边距:").grid(row=1, column=2, sticky=tk.W); ttk.Entry(crop_size_frame, textvariable=self.left_margin_var).grid(row=1, column=3, sticky=tk.EW)

        other_params_frame = ttk.LabelFrame(parent_frame, text="内容检测: 其他参数", padding="10")
        other_params_frame.pack(fill=tk.X, pady=5); other_params_frame.columnconfigure(1, weight=1); other_params_frame.columnconfigure(3, weight=1)
        ttk.Checkbutton(other_params_frame, text="清除边缘(px):", variable=self.clear_edges_var).grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(other_params_frame, textvariable=self.edge_width_var, width=8).grid(row=0, column=1, sticky=tk.EW)
        ttk.Label(other_params_frame, text="选区扩展(px):").grid(row=0, column=2, sticky=tk.E); ttk.Entry(other_params_frame, textvariable=self.expansion_var, width=8).grid(row=0, column=3, sticky=tk.EW, padx=5)
        ttk.Label(other_params_frame, text="最小面积(‱):").grid(row=1, column=0, sticky=tk.W); ttk.Entry(other_params_frame, textvariable=self.min_area_ratio_var, width=8).grid(row=1, column=1, sticky=tk.EW)
        ttk.Checkbutton(other_params_frame, text="排除宽高比>", variable=self.exclude_aspect_ratio_var).grid(row=1, column=2, sticky=tk.E); ttk.Entry(other_params_frame, textvariable=self.aspect_ratio_threshold_var, width=8).grid(row=1, column=3, sticky=tk.EW, padx=5)

        output_frame = ttk.LabelFrame(parent_frame, text="输出与调试", padding="10")
        output_frame.pack(fill=tk.X, pady=5); output_frame.columnconfigure(1, weight=1)
        ttk.Radiobutton(output_frame, text="保存到原目录", variable=self.save_option_var, value="same", command=self.toggle_output_path).grid(row=0, column=0, columnspan=3, sticky=tk.W)
        ttk.Radiobutton(output_frame, text="保存到指定目录", variable=self.save_option_var, value="specific", command=self.toggle_output_path).grid(row=1, column=0, sticky=tk.W)
        self.output_path_entry_preview = ttk.Entry(output_frame, textvariable=self.output_dir_var)
        self.output_path_entry_preview.grid(row=1, column=1, sticky=tk.EW, padx=5)
        self.btn_browse_output_preview = ttk.Button(output_frame, text="...", command=self.browse_output_dir, width=4)
        self.btn_browse_output_preview.grid(row=1, column=2, sticky=tk.E)
        ttk.Checkbutton(output_frame, text="生成调试图片", variable=self.debug_mode_var).grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
        self.toggle_output_path()

        action_button_frame = ttk.Frame(parent_frame, padding=(0, 10))
        action_button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        self.apply_btn_preview = tk.Button(action_button_frame, text="应用参数", command=self._apply_preview,
                              bg="#0078D7", fg="white", relief=tk.FLAT, padx=10, pady=5)
        self.apply_btn_preview.pack(side=tk.RIGHT)

    # --- 新增：参数配置加载/保存的UI交互函数 ---
    def _save_parameter_profile(self):
        if not self.preview_window or not self.preview_window.winfo_exists(): return
        
        config_name = simpledialog.askstring("保存配置", "请输入配置名称:", parent=self.preview_window)
        if not config_name:
            self.update_status("保存已取消。")
            return
            
        save_dir = filedialog.askdirectory(
            title="选择配置保存目录",
            initialdir=os.getcwd(), # 默认为程序当前目录
            parent=self.preview_window
        )
        if not save_dir:
            self.update_status("保存已取消。")
            return

        file_path = os.path.join(save_dir, f"{config_name.strip()}.json")
        params_dict = self._get_parameters_as_dict()
        
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(params_dict, f, indent=4)
            messagebox.showinfo("成功", f"参数配置已保存至:\n{file_path}", parent=self.preview_window)
            self.update_status(f"参数配置 '{config_name}' 已保存。")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置文件失败: {e}", parent=self.preview_window)
            self.update_status("保存配置文件失败。")

    def _load_parameter_profile(self):
        if not self.preview_window or not self.preview_window.winfo_exists(): return

        file_path = filedialog.askopenfilename(
            title="加载参数配置",
            filetypes=[("JSON config files", "*.json"), ("All files", "*.*")],
            parent=self.preview_window
        )
        if not file_path:
            self.update_status("加载已取消。")
            return
            
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            self._apply_parameters_from_dict(config)
            # 加载后自动应用预览
            self._apply_preview()
            
        except Exception as e:
            messagebox.showerror("错误", f"加载配置文件失败: {e}", parent=self.preview_window)
            self.update_status("加载配置文件失败。")

    def _on_preview_close(self):
        if self.preview_window and self.preview_window.winfo_exists():
            self.preview_window.destroy()
        self.preview_window = None

    def _set_active_slot_from_preview(self, index):
        prev_active_index = self.active_swatch_index
        self.active_swatch_index = index
        if prev_active_index != -1: self._update_all_slot_uis(prev_active_index)
        self._update_all_slot_uis(index)
        self.update_status(f"已激活颜色槽 {index + 1}，可在预览图中使用吸管取色。")

    def _on_tolerance_change(self, index):
        self._update_all_slot_uis(int(index))
        self.update_status(f"背景 {int(index)+1} 容差已修改。请点击 '应用参数' 查看效果。")

    def _reset_view(self, *args):
        if self.original_preview_image is not None and self.preview_window and self.preview_window.winfo_exists():
            # 延迟应用以确保窗口完全渲染
            self.preview_window.after(50, self._apply_preview)

    def _zoom_in(self):
        new_zoom = self.zoom_var.get() * 1.2
        self.zoom_var.set(min(8.0, new_zoom))
        self._on_zoom_scale_change(None)

    def _zoom_out(self):
        new_zoom = self.zoom_var.get() / 1.2
        self.zoom_var.set(max(0.1, new_zoom))
        self._on_zoom_scale_change(None)

    def _on_zoom_scale_change(self, value):
        if abs(self.zoom_level - self.zoom_var.get()) < 0.01: return
        self.zoom_level = self.zoom_var.get()
        self._update_zoom_label()
        if self.current_preview_cv_image is not None:
            self._update_canvas_image(self.current_preview_cv_image)

    def _update_zoom_label(self):
        self.zoom_label_var.set(f"{self.zoom_level:.0%}")

    def _on_zoom(self, event):
        if event.state & 0x4: # Ctrl Key
             if event.delta > 0 or event.num == 4: self._zoom_in()
             else: self._zoom_out()
             return "break"
        delta = -1 if (event.num == 4 or event.delta > 0) else 1
        if event.state & 0x1: self.canvas.xview_scroll(delta, "units")
        else: self.canvas.yview_scroll(delta, "units")
        return "break"

    def _on_pan_start(self, event): self.canvas.scan_mark(event.x, event.y)
    def _on_pan_drag(self, event): self.canvas.scan_dragto(event.x, event.y, gain=1)

    def _update_canvas_image(self, cv_image):
        if cv_image is None or not (self.preview_window and self.preview_window.winfo_exists()): return
        self.zoom_level = self.zoom_var.get()
        self._update_zoom_label()
        
        h, w = cv_image.shape[:2]
        new_w, new_h = int(w * self.zoom_level), int(h * self.zoom_level)
        inter_method = cv2.INTER_AREA if self.zoom_level < 1 else cv2.INTER_LINEAR
        resized_view = cv2.resize(cv_image, (new_w, new_h), interpolation=inter_method)
        img_rgb = cv2.cvtColor(resized_view, cv2.COLOR_BGR2RGB)
        self.preview_photo_image = ImageTk.PhotoImage(Image.fromarray(img_rgb))
        if self.image_on_canvas: self.canvas.delete(self.image_on_canvas)
        self.image_on_canvas = self.canvas.create_image(0, 0, anchor=tk.NW, image=self.preview_photo_image)
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def _canvas_to_image_coords(self, canvas_x, canvas_y):
        img_x = (self.canvas.canvasx(0) + canvas_x) / self.zoom_level
        img_y = (self.canvas.canvasy(0) + canvas_y) / self.zoom_level
        return int(img_x), int(img_y)

    def _on_preview_press(self, event):
        # 按下时显示原图
        self._update_canvas_image(self.original_preview_image)
        
    def _on_preview_release_and_pick(self, event):
        # 抬起时恢复效果图
        if self.original_preview_image is None: return
        
        if self.current_preview_cv_image is not None:
            self._update_canvas_image(self.current_preview_cv_image)
        else: # 如果还没有效果图，则恢复原图
            self._update_canvas_image(self.original_preview_image)

        # 执行取色
        img_x, img_y = self._canvas_to_image_coords(event.x, event.y)
        h, w = self.original_preview_image.shape[:2]
        if not (0 <= img_x < w and 0 <= img_y < h): return
        
        picked_color_bgr = self.original_preview_image[img_y, img_x]
        target_slot = self.active_swatch_index
        if target_slot == -1:
            try: target_slot = [c is None for c in self.bg_colors].index(True)
            except ValueError:
                self.update_status("所有背景色槽已满，请先激活一个槽位。"); return
        self.bg_colors[target_slot] = picked_color_bgr
        self.bg_enabled_vars[target_slot].set(True)
        self._set_active_slot_from_preview(target_slot)
        self.update_status(f"新背景色添加至槽 {target_slot+1}。请点击 '应用参数' 查看效果。")

    def _apply_preview(self):
        if not (self.preview_window and self.preview_window.winfo_exists() and hasattr(self, 'original_preview_image') and self.original_preview_image is not None):
            self.update_status("预览失败: 预览窗口或图像不存在。"); return

        # 禁用按钮，防止处理过程中被点击
        self.btn_prev.config(state="disabled")
        self.btn_next.config(state="disabled")
        self.apply_btn_preview.config(state="disabled")
        self.update_status("正在应用参数生成预览...")
        self.preview_window.update_idletasks() # 强制UI更新
        
        try:
            # 使用副本进行所有处理
            preview_image = self.original_preview_image.copy()
            
            # 调用核心处理逻辑，但仅用于生成可视化结果
            display_image, _ = self._generate_visual_steps(preview_image, return_final_preview=True)

            if display_image is not None:
                self.current_preview_cv_image = display_image
                self._update_canvas_image(self.current_preview_cv_image)
                self.update_status("预览生成完毕。蓝色:清除边缘, 红色:检测内容, 绿色:裁剪区域。")
            else:
                self.update_status("预览失败，请检查参数。")
        finally:
            # 无论成功或失败，都重新启用按钮
            self._update_nav_buttons_state() # 根据当前索引决定上一幅/下一幅的状态
            self.apply_btn_preview.config(state="normal")
            
    def _get_processing_params(self):
        # ... (此函数保持不变)
        try:
            params = {
                "crop_w": int(self.crop_width_var.get()), "crop_h": int(self.crop_height_var.get()),
                "offset_y": int(self.top_offset_var.get()), "left_margin": int(self.left_margin_var.get()),
                "expansion": int(self.expansion_var.get()), "aspect_ratio_limit": float(self.aspect_ratio_threshold_var.get()),
                "edge_width": int(self.edge_width_var.get()), "min_area_ratio": float(self.min_area_ratio_var.get())
            }
            active_colors = []
            for i in range(self.NUM_BG_COLORS):
                if self.bg_colors[i] is not None and self.bg_enabled_vars[i].get():
                    active_colors.append({"color": self.bg_colors[i], "tolerance": self.bg_tolerance_vars[i].get()})
            if not active_colors:
                messagebox.showwarning("警告", "没有启用任何背景色，无法进行内容检测。")
                return None, None
            return params, active_colors
        except (ValueError, tk.TclError) as e:
            messagebox.showerror("参数错误", f"处理失败，参数无效: {e}")
            return None, None

    def _generate_visual_steps(self, original_image, return_final_preview=False):
        # ... (此函数保持不变)
        params, active_colors = self._get_processing_params()
        if params is None:
            return (None, None) if return_final_preview else []

        steps = []
        img_copy = original_image.copy()
        steps.append(("01_Original", img_copy.copy()))

        working_image, _ = self._deskew_with_projection_profile(img_copy)
        steps.append(("02_Rotated", working_image.copy()))
        
        display_image = working_image.copy()
        overlay = np.zeros_like(display_image, dtype=np.uint8)
        alpha = 0.3
        
        img_h, img_w = working_image.shape[:2]
        significant_contours = []

        if self.clear_edges_var.get() and params["edge_width"] > 0:
            ew = min(params["edge_width"], img_h // 2, img_w // 2)
            color_bgr = active_colors[0]['color'].tolist()
            cv2.rectangle(working_image, (0, 0), (img_w - 1, ew - 1), color_bgr, -1); cv2.rectangle(working_image, (0, img_h - ew), (img_w - 1, img_h - 1), color_bgr, -1)
            cv2.rectangle(working_image, (0, 0), (ew - 1, img_h - 1), color_bgr, -1); cv2.rectangle(working_image, (img_w - ew, 0), (img_w - 1, img_h - 1), color_bgr, -1)
            steps.append(("03_Edges_Cleared_Internal", working_image.copy()))
            # For display
            cv2.rectangle(overlay, (0, 0), (img_w - 1, ew - 1), (255, 100, 0), -1) # Blue overlay
            cv2.rectangle(overlay, (0, img_h - ew), (img_w - 1, img_h - 1), (255, 100, 0), -1)
            cv2.rectangle(overlay, (0, 0), (ew - 1, img_h - 1), (255, 100, 0), -1)
            cv2.rectangle(overlay, (img_w - ew, 0), (img_w - 1, img_h - 1), (255, 100, 0), -1)

        total_background_mask = np.zeros((img_h, img_w), dtype=np.uint8)
        for ac in active_colors:
            lower = np.clip(ac["color"].astype(np.int16) - ac["tolerance"], 0, 255).astype(np.uint8)
            upper = np.clip(ac["color"].astype(np.int16) + ac["tolerance"], 0, 255).astype(np.uint8)
            total_background_mask = cv2.bitwise_or(total_background_mask, cv2.inRange(working_image, lower, upper))
        content_mask = cv2.bitwise_not(total_background_mask)
        steps.append(("04_Content_Mask", cv2.cvtColor(content_mask, cv2.COLOR_GRAY2BGR)))

        if params["expansion"] > 0:
            kernel = np.ones((params["expansion"], params["expansion"]), np.uint8)
            content_mask = cv2.dilate(content_mask, kernel, iterations=1)
            steps.append(("05_Dilated_Mask", cv2.cvtColor(content_mask, cv2.COLOR_GRAY2BGR)))
            
        contours, _ = cv2.findContours(content_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        min_area = (img_h * img_w) * (params["min_area_ratio"] / 10000.0)
        significant_contours = [c for c in contours if cv2.contourArea(c) >= min_area and not (self.exclude_aspect_ratio_var.get() and (w:=cv2.boundingRect(c)[2]) > 0 and (h:=cv2.boundingRect(c)[3]) > 0 and max(w/h, h/w) > params["aspect_ratio_limit"])]
        
        contour_img = display_image.copy()
        cv2.drawContours(contour_img, significant_contours, -1, (0, 0, 255), 3)
        steps.append(("06_Filtered_Contours", contour_img))

        if significant_contours:
            cv2.drawContours(overlay, significant_contours, -1, (0, 0, 255), -1) # Red overlay
            all_points = np.vstack(significant_contours)
            x_min, _, w, _ = cv2.boundingRect(all_points)
            x1 = x_min + w / 2 - params["crop_w"] / 2 if w >= params["crop_w"] * 0.9 else all_points[:, :, 0].min() - params["left_margin"]
            x1, y1 = max(0, x1), max(0, params["offset_y"])
            x2, y2 = min(img_w, x1 + params["crop_w"]), min(img_h, y1 + params["crop_h"])
            
            crop_box_img = display_image.copy()
            cv2.rectangle(crop_box_img, (int(x1), int(y1)), (int(x2), int(y2)), (0, 255, 0), max(3, int(img_w / 300)))
            steps.append(("07_Crop_Area", crop_box_img))
            
            final_cropped = working_image[int(y1):int(y2), int(x1):int(x2)]
            steps.append(("08_Final_Cropped", final_cropped))
        
        final_preview_img = cv2.addWeighted(overlay, alpha, display_image, 1 - alpha, 0)
        if significant_contours:
            cv2.rectangle(final_preview_img, (int(x1), int(y1)), (int(x2), int(y2)), (0, 255, 0), max(3, int(img_w / 300)))
        
        return (final_preview_img, steps) if return_final_preview else steps

    def _show_processing_steps(self):
        # ... (此函数保持不变)
        if not self.file_listbox.curselection():
            messagebox.showinfo("提示", "请先在列表中选择一张图片。")
            return

        image_path = self.file_listbox.get(self.file_listbox.curselection())
        try:
            with open(image_path, 'rb') as f: image_data = np.frombuffer(f.read(), np.uint8)
            original_image = cv2.imdecode(image_data, cv2.IMREAD_COLOR)
            if original_image is None: raise IOError("无法解码图像")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开图片: {e}"); return
        
        self.update_status("正在生成处理步骤图...")
        self.update_idletasks()
        
        steps_data = self._generate_visual_steps(original_image)
        
        if not steps_data:
            self.update_status("生成处理步骤失败，请检查参数。")
            return
            
        self._create_steps_viewer(steps_data)
        self.update_status("处理步骤显示窗口已打开。")
        
    def _create_steps_viewer(self, steps_data):
        # 创建一个模式窗口 (Toplevel)
        viewer = tk.Toplevel(self)
        viewer.title("处理过程可视化")
        viewer.geometry("1000x750")
        viewer.minsize(800, 600)

        # --- 内部状态变量 ---
        viewer.state = {
            "current_index": 0,
            "steps_data": steps_data,
            "photo_image": None,
            "zoom_level": 1.0,
            "zoom_var": tk.DoubleVar(value=1.0),
            "zoom_label_var": tk.StringVar(),
            "image_on_canvas": None,
            "current_cv_img": None
        }

        # --- 创建UI组件 ---
        
        # 1. 顶部标题栏
        title_frame = ttk.Frame(viewer, padding=(10, 10, 10, 0))
        title_frame.pack(fill=tk.X, side=tk.TOP)
        lbl_title = ttk.Label(title_frame, text="", font=("Helvetica", 14, "bold"), anchor=tk.CENTER)
        lbl_title.pack(fill=tk.X, expand=True)

        # 2. 图片显示区域 (Canvas)
        canvas_frame = ttk.Frame(viewer)
        # 使用pady在底部留出空间给导航按钮
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 40))
        
        canvas = tk.Canvas(canvas_frame, bg="gray50", cursor="fleur")
        v_scroll = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=canvas.yview)
        h_scroll = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=canvas.xview)
        canvas.config(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 3. 缩放控制 (放置在右侧)
        zoom_frame = tk.Frame(viewer)
        zoom_frame.place(relx=1.0, rely=0.5, anchor=tk.E, x=-25)

        btn_zoom_in = tk.Button(zoom_frame, text="+", width=2, relief=tk.FLAT)
        btn_zoom_in.pack(pady=(0, 2))
        zoom_slider = ttk.Scale(zoom_frame, from_=8.0, to=0.1, variable=viewer.state["zoom_var"], orient=tk.VERTICAL, length=150)
        zoom_slider.pack(pady=2, fill=tk.Y, expand=True)
        btn_zoom_out = tk.Button(zoom_frame, text="-", width=2, relief=tk.FLAT)
        btn_zoom_out.pack(pady=2)
        zoom_label = ttk.Label(zoom_frame, textvariable=viewer.state["zoom_label_var"], anchor=tk.CENTER)
        zoom_label.pack(pady=(2, 0))

        # 4. 底部导航控制 (新布局)
        nav_frame = tk.Frame(viewer)
        # 使用place将其放置在窗口底部中央
        nav_frame.place(relx=0.5, rely=1.0, anchor=tk.S, y=-10)

        btn_prev = tk.Button(nav_frame, text="< 上一幅", relief=tk.FLAT, font=("Helvetica", 10))
        btn_prev.pack(side=tk.LEFT, padx=15, ipady=3)
        
        btn_next = tk.Button(nav_frame, text="下一幅 >", relief=tk.FLAT, font=("Helvetica", 10))
        btn_next.pack(side=tk.LEFT, padx=15, ipady=3)

        # --- 核心逻辑函数 (作为嵌套函数) ---

        def update_zoom_label():
            viewer.state["zoom_label_var"].set(f"{viewer.state['zoom_level']:.0%}")

        def apply_zoom():
            if viewer.state['image_on_canvas'] is not None:
                update_display(reload_image=False)
        
        def zoom_in():
            new_zoom = viewer.state["zoom_var"].get() * 1.2
            viewer.state["zoom_var"].set(min(8.0, new_zoom))
            viewer.state['zoom_level'] = viewer.state["zoom_var"].get()
            update_zoom_label(); apply_zoom()

        def zoom_out():
            new_zoom = viewer.state["zoom_var"].get() / 1.2
            viewer.state["zoom_var"].set(max(0.1, new_zoom))
            viewer.state['zoom_level'] = viewer.state["zoom_var"].get()
            update_zoom_label(); apply_zoom()
        
        def on_zoom_scale_change(value):
            if abs(viewer.state['zoom_level'] - viewer.state["zoom_var"].get()) < 0.01: return
            viewer.state['zoom_level'] = viewer.state["zoom_var"].get()
            update_zoom_label(); apply_zoom()

        def on_mouse_wheel(event):
             if event.state & 0x4: # Ctrl Key
                 if event.delta > 0 or event.num == 4: zoom_in()
                 else: zoom_out()
                 return "break"
             delta = -1 if (event.num == 4 or event.delta > 0) else 1
             if event.state & 0x1: canvas.xview_scroll(delta, "units") # Shift Key
             else: canvas.yview_scroll(delta, "units")
             return "break"

        def on_pan_start(event): canvas.scan_mark(event.x, event.y)
        def on_pan_drag(event): canvas.scan_dragto(event.x, event.y, gain=1)
        
        def update_display(reload_image=True):
            index = viewer.state["current_index"]
            total_steps = len(viewer.state["steps_data"])
            
            if reload_image:
                current_title, viewer.state["current_cv_img"] = viewer.state["steps_data"][index]
                # 更新导航按钮状态和标题
                if index > 0:
                    prev_title = viewer.state["steps_data"][index - 1][0]
                    btn_prev.config(text=f"< {prev_title}", state=tk.NORMAL)
                else:
                    btn_prev.config(text="< 上一幅", state=tk.DISABLED)

                if index < total_steps - 1:
                    next_title = viewer.state["steps_data"][index + 1][0]
                    btn_next.config(text=f"{next_title} >", state=tk.NORMAL)
                else:
                    btn_next.config(text="下一幅 >", state=tk.DISABLED)

                lbl_title.config(text=f"步骤 {index + 1}/{total_steps}: {current_title}")

            current_cv_img = viewer.state.get("current_cv_img")
            if current_cv_img is None: return

            h, w = current_cv_img.shape[:2]
            zoom = viewer.state["zoom_level"]
            new_w, new_h = int(w * zoom), int(h * zoom)
            
            if new_w > 0 and new_h > 0:
                inter_method = cv2.INTER_AREA if zoom < 1.0 else cv2.INTER_LINEAR
                resized_img = cv2.resize(current_cv_img, (new_w, new_h), interpolation=inter_method)
                img_rgb = cv2.cvtColor(resized_img, cv2.COLOR_BGR2RGB)
                photo = ImageTk.PhotoImage(Image.fromarray(img_rgb))
                
                if viewer.state["image_on_canvas"]: canvas.delete(viewer.state["image_on_canvas"])
                
                viewer.state["image_on_canvas"] = canvas.create_image(0, 0, anchor=tk.NW, image=photo)
                viewer.state["photo_image"] = photo # 保持引用
                canvas.config(scrollregion=canvas.bbox("all"))

        def navigate(direction):
            new_index = viewer.state["current_index"] + direction
            if 0 <= new_index < len(viewer.state["steps_data"]):
                viewer.state["current_index"] = new_index
                update_display(reload_image=True)
        
        # --- 绑定事件 ---
        btn_prev.config(command=lambda: navigate(-1))
        btn_next.config(command=lambda: navigate(1))
        btn_zoom_in.config(command=zoom_in)
        btn_zoom_out.config(command=zoom_out)
        zoom_slider.config(command=on_zoom_scale_change)
        
        canvas.bind("<MouseWheel>", on_mouse_wheel)
        canvas.bind("<Button-4>", on_mouse_wheel); canvas.bind("<Button-5>", on_mouse_wheel)
        canvas.bind("<ButtonPress-2>", on_pan_start); canvas.bind("<B2-Motion>", on_pan_drag)
        canvas.bind("<ButtonPress-3>", on_pan_start); canvas.bind("<B3-Motion>", on_pan_drag)

        # --- 初始显示和模式化 ---
        update_zoom_label()
        viewer.after(50, update_display) 
        viewer.transient(self)
        viewer.grab_set()
        self.wait_window(viewer)

    def browse_files(self):
        # ... (此函数保持不变)
        files = filedialog.askopenfilenames(title="选择图片文件", filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp *.tiff"), ("All files", "*.*")])
        if files:
            current_files = set(self.file_listbox.get(0, tk.END))
            added_count = 0
            for f in files:
                if f not in current_files:
                    self.file_listbox.insert(tk.END, f)
                    added_count += 1
            self.update_status(f"添加了 {added_count} 个文件。")
            self._on_listbox_selection_change(None)

    def browse_directory(self):
        # ... (此函数保持不变)
        directory = filedialog.askdirectory(title="选择图片所在目录")
        if not directory: return
        current_files = set(self.file_listbox.get(0, tk.END))
        count = 0
        allowed_extensions = {".png", ".jpg", ".jpeg", ".bmp", ".tiff"}
        for filename in sorted(os.listdir(directory)):
            if os.path.splitext(filename)[1].lower() in allowed_extensions:
                full_path = os.path.join(directory, filename)
                if full_path not in current_files:
                    self.file_listbox.insert(tk.END, full_path)
                    count += 1
        self.update_status(f"从目录中添加了 {count} 个文件。")
        self._on_listbox_selection_change(None)

    def remove_selected(self):
        # ... (此函数保持不变)
        selected_indices = self.file_listbox.curselection()
        if not selected_indices: return
        for i in sorted(selected_indices, reverse=True): 
            self.file_listbox.delete(i)
        self.update_status("已删除选中项。")
        self._on_listbox_selection_change(None)

    def clear_list(self): 
        # ... (此函数保持不变)
        self.file_listbox.delete(0, tk.END)
        self.update_status("列表已清空。")
        self._on_listbox_selection_change(None)

    def toggle_output_path(self):
        # 在预览窗口中启用或禁用
        is_specific = self.save_option_var.get() == "specific"
        state = "normal" if is_specific else "disabled"
        if self.preview_window and self.preview_window.winfo_exists():
            self.output_path_entry_preview.config(state=state)
            self.btn_browse_output_preview.config(state=state)

    def browse_output_dir(self):
        # ... (此函数保持不变)
        directory = filedialog.askdirectory(title="选择保存目录");
        if directory: self.output_dir_var.set(directory)

    def update_status(self, message): 
        # ... (此函数保持不变)
        self.status_var.set(message); self.update_idletasks()

    def _on_closing(self):
        # ... (此函数保持不变)
        self._on_preview_close()
        self._save_config(); self.destroy()

    def _save_image_robust(self, file_path, image_data, extension=".png"):
        # ... (此函数保持不变)
        try:
            if not file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff')):
                file_ext = extension
            else:
                file_ext = os.path.splitext(file_path)[1]
            is_success, buffer = cv2.imencode(file_ext, image_data)
            if is_success:
                with open(file_path, "wb") as f: f.write(buffer)
            else: self._log(f"  > 警告: 无法编码图片 {os.path.basename(file_path)}", to_status=False)
        except Exception as e: self._log(f"  > 警告: 保存图片失败 {file_path}: {e}", to_status=False)

    def _on_listbox_select(self, event):
        # 绑定到双击事件
        if not self.file_listbox.curselection(): return
        self._open_preview_window()
        
    def _on_listbox_selection_change(self, event):
        """当列表框选择项改变时，更新“设置”按钮的状态。"""
        if self.file_listbox.curselection():
            self.btn_settings.config(state=tk.NORMAL)
        else:
            self.btn_settings.config(state=tk.DISABLED)

if __name__ == '__main__':
    app = BookCropperApp()
    app.mainloop()