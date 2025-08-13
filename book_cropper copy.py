# 扫描图片处理：将图片中的内容进行智能裁剪，并保存
import cv2
import numpy as np
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
from PIL import Image, ImageTk

class BookCropperApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # --- 实例变量 ---
        self.CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".book_cropper_config_v6.4.json")
        self.title("书籍页面智能裁剪工具 v6.4")
        self.geometry("650x750")

        # --- 参数变量 ---
        self.crop_width_var = tk.StringVar(value="2019")
        self.crop_height_var = tk.StringVar(value="3018")
        self.top_offset_var = tk.StringVar(value="10")
        self.left_margin_var = tk.StringVar(value="165")
        
        self.save_option_var = tk.StringVar(value="same")
        self.output_dir_var = tk.StringVar()
        self.debug_mode_var = tk.BooleanVar(value=True)

        # --- 内容检测参数 ---
        self.NUM_BG_COLORS = 5
        self.bg_colors = [None] * self.NUM_BG_COLORS
        self.bg_tolerance_vars = [tk.IntVar(value=25) for _ in range(self.NUM_BG_COLORS)]
        self.bg_enabled_vars = [tk.BooleanVar(value=True) for _ in range(self.NUM_BG_COLORS)]
        self.active_swatch_index = -1

        self.expansion_var = tk.StringVar(value="10")
        self.exclude_aspect_ratio_var = tk.BooleanVar(value=True)
        self.aspect_ratio_threshold_var = tk.StringVar(value="10")
        self.clear_edges_var = tk.BooleanVar(value=True)
        self.edge_width_var = tk.StringVar(value="15")
        self.min_area_ratio_var = tk.StringVar(value="5")
        
        # --- 状态与日志变量 ---
        self.preview_window = None
        self.log_messages = []
        
        # --- UI 控件引用 ---
        self.main_swatch_buttons = [] # Swatches on the main window
        self.preview_widgets = {} # Widgets inside the preview window

        self._create_widgets()
        self._load_config()
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _create_widgets(self):
        main_canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas, padding="10")

        scrollable_frame.bind("<Configure>", lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all")))

        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        file_frame = ttk.LabelFrame(scrollable_frame, text="1. 选择图片文件", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        ttk.Button(file_frame, text="浏览文件", command=self.browse_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="浏览目录", command=self.browse_directory).pack(side=tk.LEFT, padx=5)

        list_frame = ttk.LabelFrame(scrollable_frame, text="2. 待处理图片列表 (单击一项预览)", padding="10")
        list_frame.pack(fill=tk.X, pady=5)
        self.file_listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, height=6)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.file_listbox.bind("<<ListboxSelect>>", self._on_listbox_select)
        list_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=list_scrollbar.set)
        list_btn_frame = ttk.Frame(list_frame)
        list_btn_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Button(list_btn_frame, text="删除", command=self.remove_selected).pack(pady=2)
        ttk.Button(list_btn_frame, text="清空", command=self.clear_list).pack(pady=2)

        params_frame = ttk.LabelFrame(scrollable_frame, text="3. 处理参数设置", padding="10")
        params_frame.pack(fill=tk.X, pady=5)
        
        bg_frame = ttk.LabelFrame(params_frame, text="内容检测: 背景色 (在预览窗口中设置)", padding=5)
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
        ttk.Checkbutton(other_params_frame, text="清除边缘(px):", variable=self.clear_edges_var).grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(other_params_frame, textvariable=self.edge_width_var, width=8).grid(row=0, column=1, sticky=tk.EW)
        ttk.Label(other_params_frame, text="选区扩展(px):").grid(row=0, column=2, sticky=tk.E)
        ttk.Entry(other_params_frame, textvariable=self.expansion_var, width=8).grid(row=0, column=3, sticky=tk.EW, padx=5)
        ttk.Label(other_params_frame, text="最小面积(‱):").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(other_params_frame, textvariable=self.min_area_ratio_var, width=8).grid(row=1, column=1, sticky=tk.EW)
        ttk.Checkbutton(other_params_frame, text="排除宽高比>", variable=self.exclude_aspect_ratio_var).grid(row=1, column=2, sticky=tk.E)
        ttk.Entry(other_params_frame, textvariable=self.aspect_ratio_threshold_var, width=8).grid(row=1, column=3, sticky=tk.EW, padx=5)

        crop_size_frame = ttk.LabelFrame(params_frame, text="最终裁剪参数 (单位: 像素)", padding="10")
        crop_size_frame.pack(fill=tk.X, pady=2)
        crop_size_frame.columnconfigure(1, weight=1); crop_size_frame.columnconfigure(3, weight=1)
        ttk.Label(crop_size_frame, text="宽度:").grid(row=0, column=0, sticky=tk.W); ttk.Entry(crop_size_frame, textvariable=self.crop_width_var).grid(row=0, column=1, sticky=tk.EW, padx=(0,10))
        ttk.Label(crop_size_frame, text="高度:").grid(row=0, column=2, sticky=tk.W); ttk.Entry(crop_size_frame, textvariable=self.crop_height_var).grid(row=0, column=3, sticky=tk.EW)
        ttk.Label(crop_size_frame, text="顶部偏移:").grid(row=1, column=0, sticky=tk.W); ttk.Entry(crop_size_frame, textvariable=self.top_offset_var).grid(row=1, column=1, sticky=tk.EW, padx=(0,10))
        ttk.Label(crop_size_frame, text="左边距:").grid(row=1, column=2, sticky=tk.W); ttk.Entry(crop_size_frame, textvariable=self.left_margin_var).grid(row=1, column=3, sticky=tk.EW)
        
        output_frame = ttk.LabelFrame(scrollable_frame, text="4. 输出与调试", padding="10")
        output_frame.pack(fill=tk.X, pady=5)
        output_frame.columnconfigure(1, weight=1)
        ttk.Radiobutton(output_frame, text="保存到原目录", variable=self.save_option_var, value="same", command=self.toggle_output_path).grid(row=0, column=0, columnspan=3, sticky=tk.W)
        ttk.Radiobutton(output_frame, text="保存到指定目录", variable=self.save_option_var, value="specific", command=self.toggle_output_path).grid(row=1, column=0, sticky=tk.W)
        self.output_path_entry = ttk.Entry(output_frame, textvariable=self.output_dir_var, state="disabled")
        self.output_path_entry.grid(row=1, column=1, sticky=tk.EW, padx=5)
        self.btn_browse_output = ttk.Button(output_frame, text="浏览...", command=self.browse_output_dir, state="disabled")
        self.btn_browse_output.grid(row=1, column=2, sticky=tk.E)
        ttk.Checkbutton(output_frame, text="生成调试图片", variable=self.debug_mode_var).grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)

        action_frame = ttk.Frame(scrollable_frame, padding="5")
        action_frame.pack(fill=tk.X, pady=5)
        action_frame.columnconfigure(0, weight=1); action_frame.columnconfigure(1, weight=1)
        self.btn_process = ttk.Button(action_frame, text="开始处理", command=self.start_processing)
        self.btn_process.grid(row=0, column=0, sticky=tk.EW, padx=5, ipady=5)
        self.btn_log = ttk.Button(action_frame, text="查看处理日志", command=self._show_log_window)
        self.btn_log.grid(row=0, column=1, sticky=tk.EW, padx=5, ipady=5)
        
        self.status_var = tk.StringVar(value="欢迎使用！请先添加图片。")
        status_label = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=5)
        status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def _deskew_with_hough_lines(self, img):
        """
        使用霍夫变换检测图像中的直线，计算倾斜角度并进行校正。
        主要针对由文本行构成的近水平线。
        返回: (校正后的图像, 计算出的倾斜角度)
        """
        img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        # 使用Canny算子检测边缘
        img_edges = cv2.Canny(img_gray, 50, 150, apertureSize=3)
        
        # 使用霍夫变换检测直线
        # 最后一个参数是阈值，即一条直线被检测到所需的最小交点数。可以根据图像质量调整。
        lines = cv2.HoughLines(img_edges, 1, np.pi / 180, 200)

        if lines is None:
            self._log("  > 警告: Hough变换未找到任何直线，跳过旋转校正。", to_status=False)
            return img, 0.0

        angles = []
        for line in lines:
            rho, theta = line[0]
            # 霍夫变换返回的theta是法线与x轴的夹角，范围[0, pi)
            # 对于水平线，theta约为pi/2 (90度)。
            # 我们只关心接近水平的线（由文本构成）
            angle_deg = np.rad2deg(theta)
            if 45 < angle_deg < 135:  # 只考虑±45度范围内的水平线
                # 倾斜角度 = 实际角度 - 目标角度(90)
                skew_angle = angle_deg - 90
                angles.append(skew_angle)
        
        if not angles:
            self._log("  > 警告: 未在±45°范围内找到足够多的水平线，跳过旋转校正。", to_status=False)
            return img, 0.0
        
        # 使用中位数来获得最可靠的倾斜角度，以抵抗噪声
        median_angle = np.median(angles)
        
        # 避免对几乎没有倾斜或检测错误的图像进行过度旋转
        if abs(median_angle) > 45:
             self._log(f"  > 警告: 计算出的倾斜角度 {median_angle:.2f}° 过大，可能检测错误，跳过旋转。", to_status=False)
             return img, 0.0

        # 如果倾斜角度很小，则无需旋转
        if abs(median_angle) < 0.5:
             self._log(f"  > 倾斜角度 {median_angle:.2f}° 过小，跳过旋转。", to_status=False)
             return img, 0.0

        self._log(f"  > 通过Hough变换检测到倾斜角: {median_angle:.2f}°，进行校正。", to_status=False)
        h, w = img.shape[:2]
        center = (w // 2, h // 2)
        M = cv2.getRotationMatrix2D(center, median_angle, 1.0)
        
        # 使用第一个背景色作为旋转后的填充色，如果没有则用白色
        fill_color = self.bg_colors[0].tolist() if self.bg_colors[0] is not None else [255, 255, 255]

        rotated_img = cv2.warpAffine(img, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_CONSTANT, borderValue=fill_color)
        
        return rotated_img, median_angle

    def start_processing(self):
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
                
                # --- 调试步骤 1: 保存原始图片 ---
                if debug_base_path: self._save_image_robust(f"{debug_base_path}_01_original.png", original_image)
                
                # --- 步骤 1: 图像旋转校正 ---
                working_image, angle = self._deskew_with_projection_profile(original_image)
                
                # --- 调试步骤 2: 保存旋转校正后的图片 ---
                # 修改点：无论角度多小，只要开启调试模式就保存，以便确认此步骤已执行
                if debug_base_path:
                    self._save_image_robust(f"{debug_base_path}_02_rotated.png", working_image)
                
                img_h, img_w = working_image.shape[:2]

                # --- 步骤 2: 清除边缘 (可选) ---
                if self.clear_edges_var.get() and params["edge_width"] > 0:
                    ew = min(params["edge_width"], img_h // 2, img_w // 2); color_bgr = active_colors[0]['color'].tolist()
                    cv2.rectangle(working_image, (0, 0), (img_w - 1, ew - 1), color_bgr, -1); cv2.rectangle(working_image, (0, img_h - ew), (img_w - 1, img_h - 1), color_bgr, -1); cv2.rectangle(working_image, (0, 0), (ew - 1, img_h - 1), color_bgr, -1); cv2.rectangle(working_image, (img_w - ew, 0), (img_w - 1, img_h - 1), color_bgr, -1)
                    # --- 调试步骤 3: 保存清除边缘后的图片 ---
                    if debug_base_path: self._save_image_robust(f"{debug_base_path}_03_edges_cleared.png", working_image)
                
                # --- 步骤 3: 生成内容掩码 ---
                total_background_mask = np.zeros((img_h, img_w), dtype=np.uint8)
                for ac in active_colors:
                    lower = np.clip(ac["color"].astype(np.int16) - ac["tolerance"], 0, 255).astype(np.uint8)
                    upper = np.clip(ac["color"].astype(np.int16) + ac["tolerance"], 0, 255).astype(np.uint8)
                    total_background_mask = cv2.bitwise_or(total_background_mask, cv2.inRange(working_image, lower, upper))
                content_mask = cv2.bitwise_not(total_background_mask)
                # --- 调试步骤 4: 保存内容掩码 ---
                if debug_base_path: self._save_image_robust(f"{debug_base_path}_04_content_mask.png", content_mask)
                
                # --- 步骤 4: 扩展内容区域 (可选) ---
                if params["expansion"] > 0:
                    kernel = np.ones((params["expansion"], params["expansion"]), np.uint8)
                    content_mask = cv2.dilate(content_mask, kernel, iterations=1)
                    # --- 调试步骤 5: 保存扩展后的掩码 ---
                    if debug_base_path: self._save_image_robust(f"{debug_base_path}_05_dilated_mask.png", content_mask)

                # --- 步骤 5: 查找并筛选轮廓 ---
                contours, _ = cv2.findContours(content_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                min_area = (img_h * img_w) * (params["min_area_ratio"] / 10000.0)
                significant_contours = []
                for c in contours:
                    if cv2.contourArea(c) < min_area: continue
                    if self.exclude_aspect_ratio_var.get():
                        x, y, w, h = cv2.boundingRect(c)
                        if w == 0 or h == 0 or max(w/h, h/w) > params["aspect_ratio_limit"]: continue
                    significant_contours.append(c)
                
                # --- 调试步骤 6: 保存筛选后的轮廓示意图 ---
                if debug_base_path:
                    dbg_img_contours = working_image.copy()
                    cv2.drawContours(dbg_img_contours, significant_contours, -1, (0, 0, 255), 3) # 红色轮廓
                    self._save_image_robust(f"{debug_base_path}_06_filtered_contours.png", dbg_img_contours)

                if not significant_contours: self._log(f"  > 警告: 未找到有效内容。"); continue

                # --- 步骤 6: 计算裁剪区域 (含智能居中) ---
                all_points = np.vstack(significant_contours)
                content_box_x_min, _, content_width, _ = cv2.boundingRect(all_points)
                
                if content_width >= params["crop_w"] * 0.9:
                    self._log(f"  > 内容宽度({content_width}px)较大，采用居中策略。", to_status=False)
                    content_center_x = content_box_x_min + content_width / 2
                    x1 = content_center_x - (params["crop_w"] / 2)
                else:
                    self._log(f"  > 采用左边距策略 (边距: {params['left_margin']}px)。", to_status=False)
                    x1 = all_points[:, :, 0].min() - params["left_margin"]
                
                x1 = max(0, x1)
                y1 = max(0, params["offset_y"])
                x2 = min(img_w, x1 + params["crop_w"])
                y2 = min(img_h, y1 + params["crop_h"])
                
                # --- 调试步骤 7: 保存裁剪区域示意图 ---
                if debug_base_path:
                    dbg_img_crop_box = working_image.copy()
                    cv2.rectangle(dbg_img_crop_box, (int(x1), int(y1)), (int(x2), int(y2)), (0, 255, 0), 5) # 绿色裁剪框
                    self._save_image_robust(f"{debug_base_path}_07_crop_area.png", dbg_img_crop_box)

                # --- 步骤 7: 最终裁剪并保存 ---
                cropped_image = working_image[int(y1):int(y2), int(x1):int(x2)]

                # --- 调试步骤 8: 在调试目录中也保存一份最终裁剪结果 ---
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
        """
        使用投影剖面法计算图像的倾斜角度并进行校正。
        这对于文本类图像通常比Hough变换更鲁棒。
        返回: (校正后的图像, 计算出的倾斜角度)
        """
        self._log("  > 使用投影剖面法进行倾斜校正...", to_status=False)
        img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # 1. 二值化：使用Otsu's方法自动确定阈值，这对于不同亮度的扫描件很有效
        # THRESH_BINARY_INV 表示黑字（高灰度值）变白（255），白底（低灰度值）变黑（0）
        # 我们需要的是黑字为1/255，所以用INV
        _, img_binary = cv2.threshold(img_gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        
        h, w = img_binary.shape[:2]
        
        # 2. 确定搜索范围和步长
        angle_range = 5  # 搜索±5度的范围
        angle_step = 0.1 # 步长为0.1度
        angles = np.arange(-angle_range, angle_range + angle_step, angle_step)
        
        max_score = -1.0
        best_angle = 0.0
        
        # 3. 旋转迭代与评分
        for angle in angles:
            # 获取旋转矩阵
            center = (w // 2, h // 2)
            M = cv2.getRotationMatrix2D(center, angle, 1.0)
            
            # 对二值图像进行旋转
            # borderValue=0 表示旋转后空白区域填充为黑色（符合我们投影计算的需要）
            rotated_binary = cv2.warpAffine(img_binary, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_CONSTANT, borderValue=0)
            
            # 4. 计算水平投影和方差
            horizontal_projection = np.sum(rotated_binary, axis=1)
            score = np.var(horizontal_projection) # 使用方差作为评分
            
            if score > max_score:
                max_score = score
                best_angle = angle
        
        self._log(f"  > 投影法检测到最佳倾斜角: {best_angle:.2f}°", to_status=False)

        # 避免对几乎没有倾斜或检测错误的图像进行过度旋转
        if abs(best_angle) < 0.1: # 角度太小，忽略
             self._log(f"  > 倾斜角度 {best_angle:.2f}° 过小，跳过旋转。", to_status=False)
             return img, 0.0
        
        if abs(best_angle) >= angle_range: # 达到搜索边界，可能存在问题
             self._log(f"  > 警告: 倾斜角度达到搜索边界 ({best_angle:.2f}°), 可能校正不准确。", to_status=False)

        # 5. 对原始彩色图像进行最终旋转
        self._log(f"  > 进行校正...", to_status=False)
        center = (img.shape[1] // 2, img.shape[0] // 2)
        M_final = cv2.getRotationMatrix2D(center, best_angle, 1.0)
        
        # 使用第一个背景色作为旋转后的填充色，如果没有则用白色
        fill_color = self.bg_colors[0].tolist() if self.bg_colors[0] is not None else [255, 255, 255]
        rotated_img = cv2.warpAffine(img, M_final, (img.shape[1], img.shape[0]), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_CONSTANT, borderValue=fill_color)
        
        return rotated_img, best_angle

    # --- Canvas Preview Window (NO CHANGES BELOW THIS LINE, EXCEPT FOR VERSION & CONFIG FILE) ---
    def _update_all_slot_uis(self, index):
        color = self.bg_colors[index]
        tolerance = self.bg_tolerance_vars[index].get()
        main_btn = self.main_swatch_buttons[index]
        if color is not None:
            hex_color = f'#{color[2]:02x}{color[1]:02x}{color[0]:02x}'
            main_btn.config(bg=hex_color, text=f"T:{tolerance}", fg=self._get_contrasting_text_color(color))
        else:
            main_btn.config(bg="SystemButtonFace", text="空", fg="SystemButtonText")

        if self.preview_window and self.preview_window.winfo_exists():
            preview_btn = self.preview_widgets['swatches'][index]
            slot_frame = self.preview_widgets['frames'][index]
            if color is not None:
                hex_color = f'#{color[2]:02x}{color[1]:02x}{color[0]:02x}'
                preview_btn.config(bg=hex_color, text=f"T:{tolerance}", fg=self._get_contrasting_text_color(color))
            else:
                preview_btn.config(bg="SystemButtonFace", text="空", fg="SystemButtonText")
            is_active = (index == self.active_swatch_index)
            slot_frame.config(relief=tk.RIDGE if is_active else tk.GROOVE, borderwidth=3 if is_active else 2)

    def _get_contrasting_text_color(self, bgr_color):
        b, g, r = bgr_color
        brightness = (int(r) * 299 + int(g) * 587 + int(b) * 114) / 1000
        return 'white' if brightness < 128 else 'black'

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
        colors_config = []
        for i in range(self.NUM_BG_COLORS):
            colors_config.append({
                "color": self.bg_colors[i].tolist() if self.bg_colors[i] is not None else None,
                "tolerance": self.bg_tolerance_vars[i].get(),
                "enabled": self.bg_enabled_vars[i].get()
            })
        config = {
            "version": "6.4", # Updated version
            "crop_width": self.crop_width_var.get(), "crop_height": self.crop_height_var.get(),
            "top_offset": self.top_offset_var.get(), "left_margin": self.left_margin_var.get(),
            "save_option": self.save_option_var.get(), "output_dir": self.output_dir_var.get(),
            "debug_mode": self.debug_mode_var.get(), "background_colors_v6_2": colors_config,
            "expansion": self.expansion_var.get(), "exclude_aspect_ratio": self.exclude_aspect_ratio_var.get(),
            "aspect_ratio_threshold": self.aspect_ratio_threshold_var.get(), "clear_edges": self.clear_edges_var.get(),
            "edge_width": self.edge_width_var.get(), "min_area_ratio": self.min_area_ratio_var.get()
        }
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump(config, f, indent=4)
        except Exception as e: self._log(f"无法保存配置: {e}", to_status=False)

    def _load_config(self):
        if not os.path.exists(self.CONFIG_FILE): return
        try:
            with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f: config = json.load(f)
            self.crop_width_var.set(config.get("crop_width", "2019")); self.crop_height_var.set(config.get("crop_height", "3018"))
            self.top_offset_var.set(config.get("top_offset", "10")); self.left_margin_var.set(config.get("left_margin", "165"))
            self.save_option_var.set(config.get("save_option", "same")); self.output_dir_var.set(config.get("output_dir", ""))
            self.debug_mode_var.set(config.get("debug_mode", True))
            
            colors_config = config.get("background_colors_v6_2", config.get("background_colors", []))
            for i, c_conf in enumerate(colors_config):
                if i >= self.NUM_BG_COLORS: break
                self.bg_tolerance_vars[i].set(c_conf.get("tolerance", 25))
                self.bg_enabled_vars[i].set(c_conf.get("enabled", True))
                color_val = c_conf.get("color")
                self.bg_colors[i] = np.array(color_val, dtype=np.uint8) if color_val else None
                self._update_all_slot_uis(i)

            self.expansion_var.set(config.get("expansion", "10")); self.exclude_aspect_ratio_var.set(config.get("exclude_aspect_ratio", True))
            self.aspect_ratio_threshold_var.set(config.get("aspect_ratio_threshold", "10")); self.clear_edges_var.set(config.get("clear_edges", True))
            self.edge_width_var.set(config.get("edge_width", "15")); self.min_area_ratio_var.set(config.get("min_area_ratio", "5"))
            self.toggle_output_path(); self.update_status("已加载上次的配置。")
        except Exception as e: self._log(f"无法加载配置: {e}", to_status=True)

    def _open_preview_window(self):
        if self.preview_window and self.preview_window.winfo_exists(): self.preview_window.destroy()
        
        self.preview_window = tk.Toplevel(self)
        self.preview_window.title("预览与设置 (滚轮缩放, 中/右键平移, 左键按住看原图/松开取色)")
        self.preview_window.geometry("1000x700")

        self.original_preview_image = None
        self.preview_photo_image = None
        self.zoom_level = 1.0; self.pan_offset_x = 0; self.pan_offset_y = 0; self._pan_start_x = 0; self._pan_start_y = 0

        try:
            selected_path = self.file_listbox.get(self.file_listbox.curselection())
            with open(selected_path, 'rb') as f: image_data = np.frombuffer(f.read(), np.uint8)
            self.original_preview_image = cv2.imdecode(image_data, cv2.IMREAD_COLOR)
            if self.original_preview_image is None: raise IOError("无法解码图像")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开图片: {e}"); self.preview_window.destroy(); return

        main_pane = ttk.PanedWindow(self.preview_window, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True)

        canvas_frame = ttk.Frame(main_pane)
        self.canvas = tk.Canvas(canvas_frame, bg="gray50", cursor="crosshair")
        self.canvas.pack(expand=True, fill=tk.BOTH)
        main_pane.add(canvas_frame, weight=3)

        controls_frame = ttk.Frame(main_pane, padding=10)
        main_pane.add(controls_frame, weight=1)

        self.preview_widgets = {'frames': [], 'swatches': []}
        for i in range(self.NUM_BG_COLORS):
            slot_frame = ttk.LabelFrame(controls_frame, text=f"背景槽 {i+1}", padding=5)
            slot_frame.pack(fill=tk.X, pady=4)
            self.preview_widgets['frames'].append(slot_frame)
            
            top_row = ttk.Frame(slot_frame)
            top_row.pack(fill=tk.X)
            
            cb = ttk.Checkbutton(top_row, text="启用", variable=self.bg_enabled_vars[i], command=self._update_binary_preview)
            cb.pack(side=tk.LEFT)
            
            btn = tk.Button(top_row, text="空", width=8, relief=tk.RAISED, command=lambda idx=i: self._set_active_slot_from_preview(idx))
            btn.pack(side=tk.LEFT, padx=5)
            self.preview_widgets['swatches'].append(btn)
            
            clear_btn = ttk.Button(top_row, text="×", width=3, command=lambda idx=i: self._clear_color_slot(idx))
            clear_btn.pack(side=tk.RIGHT, padx=2)

            bottom_row = ttk.Frame(slot_frame)
            bottom_row.pack(fill=tk.X, pady=3)
            ttk.Label(bottom_row, text="容差:").pack(side=tk.LEFT)
            entry = ttk.Entry(bottom_row, textvariable=self.bg_tolerance_vars[i], width=5)
            entry.pack(side=tk.RIGHT)
            slider = ttk.Scale(bottom_row, from_=0, to=100, orient=tk.HORIZONTAL, variable=self.bg_tolerance_vars[i],
                               command=lambda val, idx=i: self._on_tolerance_change(idx))
            slider.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=5)
            entry.bind("<Return>", lambda e, idx=i: self._on_tolerance_change(idx))
        
        for i in range(self.NUM_BG_COLORS): self._update_all_slot_uis(i)
        
        self.canvas.bind("<Configure>", lambda e: self._reset_view()); self.canvas.bind("<MouseWheel>", self._on_zoom)
        self.canvas.bind("<Button-4>", self._on_zoom); self.canvas.bind("<Button-5>", self._on_zoom)
        self.canvas.bind("<ButtonPress-2>", self._on_pan_start); self.canvas.bind("<B2-Motion>", self._on_pan_drag)
        self.canvas.bind("<ButtonPress-3>", self._on_pan_start); self.canvas.bind("<B3-Motion>", self._on_pan_drag)
        self.canvas.bind("<ButtonPress-1>", self._on_preview_press)
        self.canvas.bind("<ButtonRelease-1>", self._on_preview_release_and_pick)
        self._reset_view()

    def _set_active_slot_from_preview(self, index):
        prev_active_index = self.active_swatch_index
        self.active_swatch_index = index
        if prev_active_index != -1: self._update_all_slot_uis(prev_active_index)
        self._update_all_slot_uis(index)
        self.update_status(f"已激活颜色槽 {index + 1}，可在预览图中使用吸管取色。")

    def _on_tolerance_change(self, index):
        self._update_all_slot_uis(int(index))
        self._update_binary_preview()

    def _clear_color_slot(self, index):
        self.bg_colors[index] = None
        self.bg_tolerance_vars[index].set(25)
        if self.active_swatch_index == index: self.active_swatch_index = -1
        self._update_all_slot_uis(index)
        self._update_binary_preview()
        self.update_status(f"已清除颜色槽 {index + 1}。")

    def _reset_view(self):
        if self.original_preview_image is None or not self.canvas.winfo_width(): return
        canvas_w, canvas_h = self.canvas.winfo_width(), self.canvas.winfo_height()
        if canvas_w == 0 or canvas_h == 0: return
        img_h, img_w = self.original_preview_image.shape[:2]
        self.zoom_level = min(canvas_w / img_w, canvas_h / img_h)
        self.pan_offset_x = (img_w - canvas_w / self.zoom_level) / 2
        self.pan_offset_y = (img_h - canvas_h / self.zoom_level) / 2
        self._update_binary_preview()
    
    def _on_zoom(self, event):
        factor = 1.1 if (event.num == 4 or event.delta > 0) else 0.9
        mouse_x, mouse_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        img_x, img_y = self._canvas_to_image_coords(mouse_x, mouse_y)
        self.pan_offset_x += (img_x - self.pan_offset_x) * (1 - 1/factor)
        self.pan_offset_y += (img_y - self.pan_offset_y) * (1 - 1/factor)
        self.zoom_level *= factor
        self._update_binary_preview()

    def _on_pan_start(self, event): self._pan_start_x, self._pan_start_y = event.x, event.y
    def _on_pan_drag(self, event):
        self.pan_offset_x -= (event.x - self._pan_start_x) / self.zoom_level
        self.pan_offset_y -= (event.y - self._pan_start_y) / self.zoom_level
        self._pan_start_x, self._pan_start_y = event.x, event.y
        self._update_binary_preview()

    def _update_canvas_image(self, cv_image):
        if cv_image is None or not (self.preview_window and self.preview_window.winfo_exists()): return
        canvas_w, canvas_h = self.canvas.winfo_width(), self.canvas.winfo_height()
        if canvas_w == 0 or canvas_h == 0: return

        M = np.float32([[self.zoom_level, 0, -self.pan_offset_x * self.zoom_level],
                        [0, self.zoom_level, -self.pan_offset_y * self.zoom_level]])
        resized_view = cv2.warpAffine(cv_image, M, (canvas_w, canvas_h), flags=cv2.INTER_NEAREST, borderValue=(128,128,128))
        img_rgb = cv2.cvtColor(resized_view, cv2.COLOR_BGR2RGB)
        self.preview_photo_image = ImageTk.PhotoImage(Image.fromarray(img_rgb))
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.preview_photo_image)

    def _canvas_to_image_coords(self, canvas_x, canvas_y):
        return (int(canvas_x / self.zoom_level + self.pan_offset_x),
                int(canvas_y / self.zoom_level + self.pan_offset_y))

    def _on_preview_press(self, event): self._update_canvas_image(self.original_preview_image)
    def _on_preview_release_and_pick(self, event):
        if self.original_preview_image is None: return
        img_x, img_y = self._canvas_to_image_coords(event.x, event.y)
        h, w = self.original_preview_image.shape[:2]
        
        if not (0 <= img_x < w and 0 <= img_y < h):
            self._update_binary_preview(); return
        
        picked_color_bgr = self.original_preview_image[img_y, img_x]
        
        target_slot = self.active_swatch_index
        if target_slot == -1:
            try: target_slot = self.bg_colors.index(None)
            except ValueError:
                self.update_status("所有背景色槽已满，请先清除或激活一个槽位。"); self._update_binary_preview(); return
        
        self.bg_colors[target_slot] = picked_color_bgr
        self.bg_enabled_vars[target_slot].set(True)
        self._set_active_slot_from_preview(target_slot)
        self.update_status(f"新背景色添加至槽 {target_slot+1}。")
        self._update_binary_preview()

    def _update_binary_preview(self, *args):
        if not (self.preview_window and self.preview_window.winfo_exists() and self.original_preview_image is not None): return
        active_colors = []
        for i in range(self.NUM_BG_COLORS):
            if self.bg_colors[i] is not None and self.bg_enabled_vars[i].get():
                active_colors.append({"color": self.bg_colors[i], "tolerance": self.bg_tolerance_vars[i].get()})
        
        if not active_colors:
            self._update_canvas_image(self.original_preview_image); return

        h, w = self.original_preview_image.shape[:2]
        total_mask = np.zeros((h, w), dtype=np.uint8)
        for ac in active_colors:
            lower = np.clip(ac["color"].astype(np.int16) - ac["tolerance"], 0, 255).astype(np.uint8)
            upper = np.clip(ac["color"].astype(np.int16) + ac["tolerance"], 0, 255).astype(np.uint8)
            total_mask = cv2.bitwise_or(total_mask, cv2.inRange(self.original_preview_image, lower, upper))
        
        binary_image = np.full_like(self.original_preview_image, (255, 255, 255), dtype=np.uint8)
        binary_image[cv2.bitwise_not(total_mask) > 0] = (0, 0, 0)
        self._update_canvas_image(binary_image)

    def browse_files(self):
        files = filedialog.askopenfilenames(title="选择图片文件", filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp *.tiff"), ("All files", "*.*")])
        if files:
            for f in files:
                if f not in self.file_listbox.get(0, tk.END): self.file_listbox.insert(tk.END, f)
            self.update_status(f"添加了 {len(files)} 个文件。")
    def browse_directory(self):
        directory = filedialog.askdirectory(title="选择图片所在目录")
        if not directory: return
        count = 0; allowed_extensions = {".png", ".jpg", ".jpeg", ".bmp", ".tiff"}
        for filename in os.listdir(directory):
            if os.path.splitext(filename)[1].lower() in allowed_extensions:
                full_path = os.path.join(directory, filename)
                if full_path not in self.file_listbox.get(0, tk.END): self.file_listbox.insert(tk.END, full_path); count += 1
        self.update_status(f"从目录中添加了 {count} 个文件。")
    def remove_selected(self):
        for i in sorted(self.file_listbox.curselection(), reverse=True): self.file_listbox.delete(i)
        self.update_status("已删除选中项。")
    def clear_list(self): self.file_listbox.delete(0, tk.END); self.update_status("列表已清空。")
    def toggle_output_path(self):
        state = "normal" if self.save_option_var.get() == "specific" else "disabled"
        self.output_path_entry.config(state=state); self.btn_browse_output.config(state=state)
    def browse_output_dir(self):
        directory = filedialog.askdirectory(title="选择保存目录");
        if directory: self.output_dir_var.set(directory)
    def update_status(self, message): self.status_var.set(message); self.update_idletasks()
    def _on_closing(self):
        if self.preview_window and self.preview_window.winfo_exists(): self.preview_window.destroy()
        self._save_config(); self.destroy()
    def _save_image_robust(self, file_path, image_data, extension=".png"):
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
        if not self.file_listbox.curselection(): return
        self._open_preview_window()

if __name__ == '__main__':
    app = BookCropperApp()
    app.mainloop()