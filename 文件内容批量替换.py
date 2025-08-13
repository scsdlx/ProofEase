import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import shutil
import threading
import re
from datetime import datetime
import locale

class BatchFileContentReplacerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("高级批量文件内容替换工具 (v3.0 - 内容替换)")
        self.geometry("1000x750") # 增加了高度以容纳新选项

        # 数据存储
        self.file_list_data = [] # 存储文件的详细信息
        self.file_counter = 0
        
        # 排序状态: 默认按文件名升序
        self.sort_column = "filename"
        self.sort_reverse = False

        self._create_widgets()
        self._update_treeview_headers() # 初始化标题排序标志

    def _create_widgets(self):
        # --- 顶部：文件操作区 ---
        top_frame = ttk.Frame(self, padding="10")
        top_frame.pack(fill=tk.X)
        ttk.Button(top_frame, text="添加文件", command=self._add_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="添加文件夹", command=self._add_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="删除选中", command=self._remove_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="清空列表", command=self._clear_list).pack(side=tk.LEFT, padx=5)

        # --- 中部：文件列表区 ---
        list_frame = ttk.Frame(self, padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)

        self.columns = ("#", "filename", "filetype", "filepath", "mtime", "size_kb")
        self.tree = ttk.Treeview(list_frame, columns=self.columns, show="headings")
        
        self.col_map = {
            "#": "序号", "filename": "文件名", "filetype": "类型", 
            "filepath": "所在目录", "mtime": "修改时间", "size_kb": "大小 (KB)"
        }

        # 为所有列标题设置文本和排序命令
        for col, text in self.col_map.items():
            self.tree.heading(col, text=text, command=lambda c=col: self._sort_column(c))
        
        self.tree.column("#", width=60, anchor="center")
        self.tree.column("filename", width=250)
        self.tree.column("filetype", width=80, anchor="center")
        self.tree.column("filepath", width=300)
        self.tree.column("mtime", width=150, anchor="center")
        self.tree.column("size_kb", width=100, anchor="e")

        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(list_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # --- 底部：设置与执行区 ---
        bottom_frame = ttk.Frame(self, padding="10")
        bottom_frame.pack(fill=tk.X)

        # 替换规则
        rule_frame = ttk.LabelFrame(bottom_frame, text="内容替换规则", padding="10")
        rule_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(rule_frame, text="待替换内容:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.find_entry = ttk.Entry(rule_frame, width=40)
        self.find_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Label(rule_frame, text="替换为:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.replace_entry = ttk.Entry(rule_frame, width=40)
        self.replace_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.use_regex_var = tk.BooleanVar()
        ttk.Checkbutton(rule_frame, text="使用通配符/特殊字符 (Regex)", variable=self.use_regex_var).grid(row=2, column=1, sticky='w', pady=5)
        ttk.Button(rule_frame, text="通配符使用方法", command=self._show_help).grid(row=2, column=2, padx=10)
        rule_frame.columnconfigure(1, weight=1)

        # 输出选项
        output_frame = ttk.LabelFrame(bottom_frame, text="输出与编码选项", padding="10")
        output_frame.pack(fill=tk.X, pady=5)
        
        # Encoding options
        encoding_frame = ttk.Frame(output_frame)
        encoding_frame.pack(fill=tk.X, anchor="w", pady=(0, 10))
        ttk.Label(encoding_frame, text="文件编码:").pack(side=tk.LEFT, padx=(0, 5))
        self.encoding_var = tk.StringVar(value="utf-8")
        # 尝试获取系统默认编码
        try:
            default_encoding = locale.getpreferredencoding()
        except Exception:
            default_encoding = "utf-8"
        encodings = ["utf-8", "gbk", "gb2312", "latin-1", default_encoding]
        self.encoding_combo = ttk.Combobox(encoding_frame, textvariable=self.encoding_var, values=list(set(encodings)), width=15)
        self.encoding_combo.pack(side=tk.LEFT)
        ttk.Label(encoding_frame, text="(请选择正确的编码读取和保存文件)").pack(side=tk.LEFT, padx=10)

        # Output options
        self.output_option = tk.StringVar(value="original")
        ttk.Radiobutton(output_frame, text="在原位置修改文件内容", variable=self.output_option, value="original", command=self._toggle_output_dir).pack(anchor="w")
        specific_dir_frame = ttk.Frame(output_frame)
        specific_dir_frame.pack(fill=tk.X, anchor="w")
        ttk.Radiobutton(specific_dir_frame, text="将修改后的文件保存到新目录", variable=self.output_option, value="specific", command=self._toggle_output_dir).pack(side=tk.LEFT)
        self.output_dir_entry = ttk.Entry(specific_dir_frame, width=50, state="disabled")
        self.output_dir_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        self.browse_button = ttk.Button(specific_dir_frame, text="浏览...", command=self._browse_output_dir, state="disabled")
        self.browse_button.pack(side=tk.LEFT)

        # 进度和状态
        progress_frame = ttk.Frame(bottom_frame, padding="5 0")
        progress_frame.pack(fill=tk.X, pady=5)
        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill=tk.X, expand=True)
        self.status_label = ttk.Label(progress_frame, text="准备就绪", anchor="w")
        self.status_label.pack(fill=tk.X, expand=True)

        # 执行按钮
        self.start_button = ttk.Button(bottom_frame, text="开始替换内容", command=self._start_processing)
        self.start_button.pack(pady=10)

    def _add_to_list(self, file_paths):
        existing_paths = {item['path'] for item in self.file_list_data}
        new_files_added = 0
        for f in file_paths:
            if f not in existing_paths and os.path.isfile(f): # 确保是文件
                try:
                    stat = os.stat(f)
                    _, extension = os.path.splitext(f)
                    file_info = {
                        "id": self.file_counter,
                        "path": f,
                        "filename": os.path.basename(f),
                        "filepath": os.path.dirname(f),
                        "filetype": extension[1:].lower() if extension else '无',
                        "mtime_ts": stat.st_mtime,
                        "mtime_str": datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                        "size_kb": round(stat.st_size / 1024, 2)
                    }
                    self.file_list_data.append(file_info)
                    self.file_counter += 1
                    new_files_added += 1
                except OSError as e:
                    print(f"无法访问文件: {f}, 错误: {e}")
        if new_files_added > 0:
            self._sort_and_refresh_view()
        self.status_label.config(text=f"添加了 {new_files_added} 个新文件。当前共 {len(self.file_list_data)} 个文件。")
    
    def _update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for item in self.file_list_data:
            self.tree.insert("", "end", iid=item['id'], values=(
                item['id'] + 1, item['filename'], item['filetype'], 
                item['filepath'], item['mtime_str'], f"{item['size_kb']:,}"
            ))
        self.status_label.config(text=f"列表已更新。当前共 {len(self.file_list_data)} 个文件。")

    def _update_treeview_headers(self):
        for col, text in self.col_map.items():
            if col == self.sort_column:
                sort_indicator = ' ▼' if self.sort_reverse else ' ▲'
                self.tree.heading(col, text=text + sort_indicator)
            else:
                self.tree.heading(col, text=text)

    def _sort_column(self, col):
        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = col
            self.sort_reverse = False
        self._sort_and_refresh_view()

    def _sort_and_refresh_view(self):
        key_map = {
            '#':        lambda item: item['id'],
            'size_kb':  lambda item: item['size_kb'],
            'mtime':    lambda item: item['mtime_ts']
        }
        sort_key = key_map.get(
            self.sort_column, 
            lambda item: str(item.get(self.sort_column, "")).lower()
        )
        self.file_list_data.sort(key=sort_key, reverse=self.sort_reverse)
        self._update_treeview_headers()
        self._update_treeview()

    def _remove_selected(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先在列表中选择要删除的文件。")
            return
        selected_ids = {int(i) for i in selected_items}
        self.file_list_data = [item for item in self.file_list_data if item['id'] not in selected_ids]
        self._update_treeview()

    def _clear_list(self):
        if not self.file_list_data:
            return
        if messagebox.askyesno("确认", "确定要清空整个文件列表吗？"):
            self.file_list_data.clear()
            self.file_counter = 0
            self._update_treeview()
            self.status_label.config(text="列表已清空。")
    
    def _show_help(self):
        help_text = """
        当“使用通配符/特殊字符 (Regex)”被勾选时，系统将使用正则表达式对【文件内容】进行匹配和替换。

        【基础概念】
        - 待替换内容：一个正则表达式模式。
        - 替换为：一个替换字符串，可以使用捕获组。

        【常用元字符 (在“待替换内容”中使用)】
        .         : 匹配除换行符外的任意单个字符
        *         : 匹配前一个字符0次或多次 (例如, "a*" 匹配 "", "a", "aa")
        +         : 匹配前一个字符1次或多次 (例如, "a+" 匹配 "a", "aa")
        \d        : 匹配任意数字 (等同于 [0-9])
        \s        : 匹配任意空白字符 (空格, 制表符 \t 等)
        ()        : 捕获组。将括号内的模式匹配到的内容捕获，可在“替换为”中引用。
        ^         : 匹配一行的开头（需要配合多行模式）
        $         : 匹配一行的结尾（需要配合多行模式）
        
        【特殊字符】
        \\t        : 制表符 (Tab)
        \\n        : 换行符 (Newline)
        \\. \\* \\( 等 : 若要匹配元字符本身，需在其前加反斜杠 \\ 转义。

        【“替换为”中的用法 - 捕获组引用】
        \\1, \\2, ... : 引用“待替换内容”中第1, 2, ...个括号捕获到的内容。
        
        【示例】
        1. 替换所有旧网址为新网址:
           - 文件内容包含: "访问 http://old-site.com 获取更多"
           - 待替换内容: http://old-site\\.com
           - 替换为: https://new-site.com
           - 结果: "访问 https://new-site.com 获取更多"
        
        2. 更新配置文件中的版本号:
           - 文件内容: version = "1.2.3"
           - 待替换内容: (version = ")\\d+\\.\\d+\\.\\d+(")
           - 替换为: \\12.0.0\\2
           - 结果: version = "2.0.0"
        
        注意：不勾选此框时，为普通文本替换。正则表达式默认作用于整个文件内容。
        """
        messagebox.showinfo("通配符/特殊字符 (Regex) 使用方法", help_text)

    def _add_files(self):
        files = filedialog.askopenfilenames(title="选择要处理的文件")
        if files:
            self._add_to_list(files)

    def _add_folder(self):
        folder = filedialog.askdirectory(title="选择包含文件的文件夹")
        if folder:
            files_to_add = []
            for root, _, filenames in os.walk(folder):
                for filename in filenames:
                    files_to_add.append(os.path.join(root, filename))
            if files_to_add:
                self._add_to_list(files_to_add)
            else:
                messagebox.showinfo("提示", "所选文件夹中没有文件。")
    
    def _toggle_output_dir(self):
        if self.output_option.get() == "specific":
            self.output_dir_entry.config(state="normal")
            self.browse_button.config(state="normal")
        else:
            self.output_dir_entry.config(state="disabled")
            self.browse_button.config(state="disabled")

    def _browse_output_dir(self):
        directory = filedialog.askdirectory(title="选择目标文件夹")
        if directory:
            self.output_dir_entry.config(state="normal")
            self.output_dir_entry.delete(0, tk.END)
            self.output_dir_entry.insert(0, directory)

    def _start_processing(self):
        if not self.file_list_data:
            messagebox.showerror("错误", "文件列表为空，请先添加文件。")
            return
        find_what = self.find_entry.get()
        if not find_what:
            messagebox.showerror("错误", "“待替换内容”不能为空。")
            return
        
        if self.use_regex_var.get():
            try:
                re.compile(find_what)
            except re.error as e:
                messagebox.showerror("正则表达式错误", f"“待替换内容”中的正则表达式无效：\n{e}")
                return
            
        replace_with = self.replace_entry.get()
        output_mode = self.output_option.get()
        output_dir = self.output_dir_entry.get()
        encoding = self.encoding_var.get()

        if output_mode == "specific" and not (output_dir and os.path.isdir(output_dir)):
            messagebox.showerror("错误", "请选择一个有效的目标目录。")
            return
        
        confirm_msg = (
            f"将在 {len(self.file_list_data)} 个文件的【内容】中查找 “{find_what}” 并替换为 “{replace_with}”。\n"
            f"模式: {'正则表达式' if self.use_regex_var.get() else '普通文本'}\n"
            f"文件编码: {encoding}\n"
            f"输出方式: {'在原位置修改' if output_mode == 'original' else f'保存到目录: {output_dir}'}\n\n"
            "此操作可能无法撤销，确定要继续吗？"
        )
        if not messagebox.askyesno("请确认操作", confirm_msg):
            return
            
        self.start_button.config(state="disabled")
        processing_thread = threading.Thread(
            target=self._process_files_thread,
            args=(find_what, replace_with, output_mode, output_dir, self.use_regex_var.get(), encoding),
            daemon=True
        )
        processing_thread.start()

    def _process_files_thread(self, find_what, replace_with, output_mode, output_dir, use_regex, encoding):
        total_files = len(self.file_list_data)
        self.progress_bar["maximum"] = total_files
        self.progress_bar["value"] = 0
        
        success_count, fail_count, skipped_count = 0, 0, 0
        files_to_process = list(self.file_list_data)
        updated_file_ids = []

        for i, file_info in enumerate(files_to_process):
            original_path = file_info['path']
            original_filename = file_info['filename']

            self.status_label.config(text=f"正在处理: {original_filename} ({i+1}/{total_files})")
            self.progress_bar["value"] = i + 1

            try:
                with open(original_path, 'r', encoding=encoding) as f:
                    content = f.read()
            except (UnicodeDecodeError, IOError) as e:
                print(f"读取文件失败 (可能编码错误): {original_path}, 错误: {e}")
                fail_count += 1
                continue
            
            new_content = ""
            content_changed = False
            if use_regex:
                new_content, num_subs = re.subn(find_what, replace_with, content, flags=re.MULTILINE)
                if num_subs > 0:
                    content_changed = True
            else:
                if find_what in content:
                    new_content = content.replace(find_what, replace_with)
                    content_changed = True

            if not content_changed:
                # 如果是“保存到新目录”模式，即使内容未变，也复制原文件
                if output_mode == "specific":
                    try:
                        dest_path = os.path.join(output_dir, original_filename)
                        if os.path.abspath(original_path) != os.path.abspath(dest_path):
                            if os.path.exists(dest_path):
                                fail_count += 1 # 目标文件已存在
                            else:
                                shutil.copy2(original_path, dest_path)
                                success_count += 1
                        else: # 源和目标相同
                            skipped_count += 1
                    except Exception as e:
                        fail_count += 1
                        print(f"复制文件失败: {original_path}, 错误: {e}")
                else: # 原位置修改模式，无变化则跳过
                    skipped_count += 1
                continue

            # --- 内容已改变，执行写入操作 ---
            try:
                if output_mode == "original":
                    # 安全写入：先写临时文件，再替换原文件
                    temp_path = original_path + ".tmpreplace"
                    with open(temp_path, 'w', encoding=encoding) as f:
                        f.write(new_content)
                    os.replace(temp_path, original_path) # 原子操作
                    updated_file_ids.append(file_info['id'])
                else: # 保存到新目录
                    dest_path = os.path.join(output_dir, original_filename)
                    if os.path.exists(dest_path):
                        fail_count += 1 # 目标文件已存在
                        continue
                    with open(dest_path, 'w', encoding=encoding) as f:
                        f.write(new_content)
                
                success_count += 1
            except Exception as e:
                fail_count += 1
                print(f"写入文件失败: {original_path}, 错误: {e}")
        
        self.after(0, self._on_processing_complete, success_count, fail_count, skipped_count, updated_file_ids)


    def _on_processing_complete(self, success, fail, skipped, updated_file_ids):
        self.start_button.config(state="normal")
        
        # 如果是原位置修改，则更新列表中的文件信息（大小、修改时间）
        if self.output_option.get() == "original" and updated_file_ids:
            self.status_label.config(text="正在更新文件信息...")
            self.update_idletasks() # 强制UI更新
            
            data_map = {item['id']: item for item in self.file_list_data}
            for file_id in updated_file_ids:
                if file_id in data_map:
                    file_info = data_map[file_id]
                    try:
                        stat = os.stat(file_info['path'])
                        file_info['mtime_ts'] = stat.st_mtime
                        file_info['mtime_str'] = datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                        file_info['size_kb'] = round(stat.st_size / 1024, 2)
                    except OSError as e:
                        print(f"更新文件信息失败: {file_info['path']}, {e}")
            self._sort_and_refresh_view() # 使用更新后的数据刷新视图
        
        self.status_label.config(text="处理完成！")
        summary_msg = (
            f"处理完成！\n\n"
            f"成功修改/复制: {success} 个\n"
            f"失败/目标已存在/编码错误: {fail} 个\n"
            f"跳过 (无内容匹配): {skipped} 个\n\n"
            f"总计处理: {len(self.file_list_data)} 个文件"
        )
        messagebox.showinfo("处理结果", summary_msg)
        self.progress_bar["value"] = 0

if __name__ == "__main__":
    app = BatchFileContentReplacerApp()
    app.mainloop()