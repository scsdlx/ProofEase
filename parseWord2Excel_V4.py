# 1. 可以浏览文件或目录，多次添加文件，并管理文件列表
# 2. 可以选择要提取的层级，并按层级，而还是样式名称提取内容
# 3. 标题及内容带自动编号
# 4. 提取内容保存为Excel文件
import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import threading
import pandas as pd
import win32com.client as win32

# --- 核心提取逻辑 (已修改为包含自动编号) ---

def clean_text(text):
    """
    使用正则表达式移除文本中非法的XML字符。
    """
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)

def extract_word_to_excel(file_list, output_dir, num_levels_to_extract, status_callback, progress_callback):
    """
    提取指定文件列表中的所有Word文档的内容到Excel表中。
    最终版逻辑 V5：
    1. 接收一个包含Word文档完整路径的列表。
    2. 通过段落的 OutlineLevel (大纲级别) 识别标题，更加可靠。
    3. 【已修改】提取标题和列表项时，会将其自动编号 (如 "第一章", "1.1", "(一)") 一同提取。
    4. 自动检测每个文档的最高标题级别，并从该级别开始提取N级。
    5. 输出的Excel列为：word文档名称, 第1层标题, ..., 第N层标题, 内容。
    6. 返回处理结果的总结信息。

    :param file_list: 存放Word文档完整路径的列表。
    :param output_dir: 输出Excel文件的目录。
    :param num_levels_to_extract: 要提取的标题层级数。
    :param status_callback: 用于向UI发送状态更新的回调函数。
    :param progress_callback: 用于向UI更新进度条的回调函数。
    :return: 一个包含总结信息的字典。
    """
    result_summary = {
        "success": False, "message": "", "output_path": "",
        "files_processed": 0, "total_files": 0, "max_level_found_overall": 0,
    }

    if not file_list:
        result_summary["message"] = "错误：文件列表为空。"
        return result_summary

    word_app = None
    try:
        status_callback("正在启动Word应用程序...")
        word_app = win32.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = False
    except Exception as e:
        result_summary["message"] = f"错误：无法启动Word应用程序，请确保已正确安装Office。\n详细错误: {e}"
        return result_summary

    result_summary["total_files"] = len(file_list)
    status_callback(f"找到 {len(file_list)} 个待处理的Word文档，准备开始处理...")

    all_data_from_docs = []
    max_level_found_overall = 0

    try:
        for i, file_path in enumerate(file_list):
            progress_callback((i + 1) / len(file_list) * 100)
            filename = os.path.basename(file_path)
            status_callback(f"正在处理: {filename} ({i+1}/{len(file_list)})")
            
            doc = None
            try:
                doc = word_app.Documents.Open(file_path)
                
                doc_content_aggregator = {}
                current_headings = {f'标题{i}': '' for i in range(1, 10)} 

                all_levels_in_doc = {p.OutlineLevel for p in doc.Paragraphs if 1 <= p.OutlineLevel <= 9}

                if not all_levels_in_doc:
                    status_callback(f"  -> 警告: 文件 '{filename}' 中未找到任何大纲级别（1-9级）的标题，已跳过。")
                    doc.Close(SaveChanges=False)
                    continue

                min_level_in_doc = min(all_levels_in_doc)
                actual_levels_found_in_doc = max(all_levels_in_doc) - min_level_in_doc + 1
                max_level_found_overall = max(max_level_found_overall, actual_levels_found_in_doc)

                for para in doc.Paragraphs:
                    # --- START OF MODIFICATION TO INCLUDE NUMBERING ---
                    raw_text = para.Range.Text
                    list_string = para.Range.ListFormat.ListString
                    
                    if list_string:
                        full_text = f"{list_string} {raw_text}"
                    else:
                        full_text = raw_text

                    para_text = clean_text(full_text).strip()
                    # --- END OF MODIFICATION ---

                    if not para_text:
                        continue
                    
                    level = para.OutlineLevel 

                    if 1 <= level <= 9:
                        current_headings[f'标题{level}'] = para_text
                        for L in range(level + 1, 10):
                            current_headings[f'标题{L}'] = ''
                    
                    key_headings = []
                    for j in range(num_levels_to_extract):
                        absolute_level_num = min_level_in_doc + j
                        if absolute_level_num <= 9:
                            heading_text = current_headings.get(f'标题{absolute_level_num}', '')
                            key_headings.append(heading_text)
                    key_tuple = tuple(key_headings)
                    
                    if key_tuple not in doc_content_aggregator:
                        doc_content_aggregator[key_tuple] = []
                    doc_content_aggregator[key_tuple].append(para_text)

                for headings_tuple, content_list in doc_content_aggregator.items():
                    if not any(headings_tuple):
                        continue
                    full_content = "\n".join(content_list).strip()
                    if len(full_content) > 200: 
                        row_data = {'word文档名称': filename}
                        for j, heading_text in enumerate(headings_tuple):
                            relative_title_key = f'第{j+1}层标题'
                            row_data[relative_title_key] = heading_text
                        row_data['内容'] = full_content
                        all_data_from_docs.append(row_data)
                    else:
                        row_data = {'word文档名称': filename}
                        for j, heading_text in enumerate(headings_tuple):
                            relative_title_key = f'第{j+1}层标题'
                            row_data[relative_title_key] = heading_text
                        row_data['内容'] = full_content
                        all_data_from_docs.append(row_data)
                    
                doc.Close(SaveChanges=False)
                result_summary["files_processed"] += 1
            
            except Exception as e:
                status_callback(f"\n处理文件 '{filename}' 时发生错误: {e}")
                if doc:
                    try: doc.Close(SaveChanges=False)
                    except: pass
                continue
    finally:
        if word_app:
            word_app.Quit()
            status_callback("Word应用程序已关闭。")

    if not all_data_from_docs:
        result_summary["message"] = "处理完毕，但未能从任何文档中提取到符合条件（内容总长度>200字符）的数据。"
        return result_summary

    status_callback("正在整合数据并写入Excel文件...")
    
    final_columns = ['word文档名称']
    for i in range(1, num_levels_to_extract + 1):
        final_columns.append(f'第{i}层标题')
    final_columns.append('内容')
    
    df = pd.DataFrame(all_data_from_docs)
    df = df.reindex(columns=final_columns).fillna('')
    
    output_filename = f"教材内容提取结果_前{num_levels_to_extract}级标题.xlsx"
    output_excel_path = os.path.join(output_dir, output_filename)
    
    try:
        df.to_excel(output_excel_path, index=False, engine='openpyxl')
        result_summary["success"] = True
        result_summary["message"] = f"处理完成！共处理 {result_summary['files_processed']} / {result_summary['total_files']} 个文件。"
        result_summary["output_path"] = os.path.abspath(output_excel_path)
        result_summary["max_level_found_overall"] = max_level_found_overall
    except Exception as e:
        result_summary["message"] = f"写入Excel文件时发生错误: {e}"

    return result_summary


# --- UI界面类 (与V4版本相同，无需修改) ---
class WordExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word 内容提取工具 V5") # 版本号更新
        self.root.geometry("700x550")
        self.result_path = None
        # ... (UI代码与之前完全相同) ...
        # --- UI组件 ---
        # 1. 文件选择按钮区
        self.selection_frame = ttk.Frame(root, padding="10")
        self.selection_frame.pack(fill=tk.X)
        ttk.Button(self.selection_frame, text="选择文件...", command=self.browse_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(self.selection_frame, text="选择文件夹...", command=self.browse_directory).pack(side=tk.LEFT)

        # 2. 文件列表显示区
        self.list_frame = ttk.Frame(root, padding=(10, 0, 10, 5))
        self.list_frame.pack(fill=tk.BOTH, expand=True)
        self.file_listbox = tk.Listbox(self.list_frame, selectmode=tk.EXTENDED)
        self.list_scrollbar = ttk.Scrollbar(self.list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.config(yscrollcommand=self.list_scrollbar.set)
        self.list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 3. 列表管理和操作区
        self.action_frame = ttk.Frame(root, padding=(10, 5, 10, 10))
        self.action_frame.pack(fill=tk.X)
        ttk.Button(self.action_frame, text="移除选中", command=self.remove_selected_files).pack(side=tk.LEFT)
        ttk.Button(self.action_frame, text="清空列表", command=self.clear_file_list).pack(side=tk.LEFT, padx=(5, 20))
        
        ttk.Label(self.action_frame, text="提取前").pack(side=tk.LEFT, padx=(0, 5))
        self.level_var = tk.StringVar(value='3')
        ttk.Combobox(self.action_frame, textvariable=self.level_var, values=['1', '2', '3', '4', '5', '6', '7', '8'], state='readonly', width=5).pack(side=tk.LEFT)
        ttk.Label(self.action_frame, text="级标题").pack(side=tk.LEFT, padx=(5, 10))

        self.start_button = ttk.Button(self.action_frame, text="开始提取", command=self.start_extraction_thread)
        self.start_button.pack(side=tk.LEFT)
        self.open_result_button = ttk.Button(self.action_frame, text="打开转换结果", command=self.open_result_file)

        # 4. 状态和进度区
        self.status_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=10, state='disabled')
        self.status_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))
        self.status_text.tag_config('error', foreground='red', font=('TkDefaultFont', 9, 'bold'))
        
        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10, pady=(0, 10))

    def browse_files(self):
        file_paths = filedialog.askopenfilenames(
            title="选择一个或多个Word文件",
            filetypes=[("Word Documents", "*.docx *.doc"), ("All files", "*.*")]
        )
        if file_paths:
            current_files = self.file_listbox.get(0, tk.END)
            for path in file_paths:
                if path not in current_files:
                    self.file_listbox.insert(tk.END, path)
            self.hide_open_result_button()

    def browse_directory(self):
        directory = filedialog.askdirectory(title="选择一个文件夹")
        if directory:
            current_files = self.file_listbox.get(0, tk.END)
            for filename in os.listdir(directory):
                if filename.endswith(('.doc', '.docx')) and not filename.startswith('~'):
                    full_path = os.path.join(directory, filename)
                    if full_path not in current_files:
                        self.file_listbox.insert(tk.END, full_path)
            self.hide_open_result_button()

    def remove_selected_files(self):
        selected_indices = self.file_listbox.curselection()
        for i in reversed(selected_indices):
            self.file_listbox.delete(i)
        self.hide_open_result_button()
            
    def clear_file_list(self):
        self.file_listbox.delete(0, tk.END)
        self.hide_open_result_button()

    def log_status(self, message):
        self.root.after(0, self._log_status_sync, message)
        
    def _log_status_sync(self, message):
        self.status_text.config(state='normal')
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')

    def update_progress(self, value):
        self.root.after(0, self.progress.config, {'value': value})
        
    def start_extraction_thread(self):
        file_list = self.file_listbox.get(0, tk.END)
        if not file_list:
            self._log_status_sync("错误：请先选择至少一个文件或文件夹。")
            return

        self.start_button.config(state='disabled')
        self.progress['value'] = 0
        self.status_text.config(state='normal')
        self.status_text.delete('1.0', tk.END)
        self.status_text.config(state='disabled')
        self.hide_open_result_button()

        num_levels = int(self.level_var.get())

        thread = threading.Thread(
            target=self.run_extraction,
            args=(list(file_list), num_levels)
        )
        thread.daemon = True
        thread.start()

    def run_extraction(self, file_list, num_levels):
        try:
            output_dir = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            output_dir = os.getcwd()
            
        result = extract_word_to_excel(
            file_list, output_dir, num_levels,
            self.log_status,
            self.update_progress
        )
        self.root.after(0, self.on_extraction_complete, result, num_levels)

    def on_extraction_complete(self, result, requested_levels):
        self._log_status_sync("\n--- 提取总结 ---")
        self._log_status_sync(result['message'])

        if result['success']:
            self._log_status_sync(f"结果已保存到: {result['output_path']}")
            self.result_path = result['output_path']
            self.show_open_result_button()

            if result['max_level_found_overall'] < requested_levels:
                warning_msg = (f"警告：您选择了提取 {requested_levels} 个层级的标题，"
                               f"但在所有文件中最多只找到了 {result['max_level_found_overall']} 个层级。\n"
                               "请检查您的Word文档标题设置是否正确。")
                self.status_text.config(state='normal')
                self.status_text.insert(tk.END, "\n" + warning_msg + "\n", 'error')
                self.status_text.see(tk.END)
                self.status_text.config(state='disabled')
        
        self.start_button.config(state='normal')
        if result['success']:
             self.progress['value'] = 100
        else:
             self.progress['value'] = 0

    def show_open_result_button(self):
        self.open_result_button.pack(side=tk.LEFT, padx=(5, 0))

    def hide_open_result_button(self):
        self.result_path = None
        self.open_result_button.pack_forget()
        
    def open_result_file(self):
        if self.result_path and os.path.exists(self.result_path):
            try:
                os.startfile(self.result_path)
            except Exception as e:
                self._log_status_sync(f"错误：无法打开文件 {self.result_path}\n{e}")
        else:
            self._log_status_sync(f"错误：结果文件不存在或路径无效。请重新提取。")

if __name__ == '__main__':
    root = tk.Tk()
    app = WordExtractorApp(root)
    root.mainloop()