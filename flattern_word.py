# 1. 可以浏览文件或目录，多次添加文件，并管理文件列表
# 2. 【已修改】需要输入教材ID(material_id)
# 3. 【已修改】将提取的标题和内容，按照层级结构存入数据库
# 4. 标题及内容带自动编号
import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import threading
from datetime import datetime
import win32com.client as win32
import pythoncom
import mysql.connector

# --- 导入数据库配置 ---
try:
    from db_config import DB_CONFIG
except ImportError:
    # 如果 db_config.py 不存在，提供一个默认的空配置，并在UI层面报错
    DB_CONFIG = {}
    print("错误: 无法导入 db_config.py 文件。请确保该文件存在且配置正确。")

# Word应用中的常量
WD_OUTLINE_LEVEL_BODY_TEXT = 10

# --- 核心提取与数据库存储逻辑 ---

def clean_text(text):
    """
    使用正则表达式移除文本中非法的XML字符，并清理Word特有的控制字符。
    """
    if not text:
        return ""
    # 移除大多数控制字符，但保留换行符等
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)
    # Word中段落末尾常有\r\x07，需要清理
    text = text.replace('\r\x07', '').replace('\x07', '')
    return text.strip()

def extract_word_and_save_to_db(file_list, material_id, status_callback, progress_callback):
    """
    提取Word文档内容并按层级结构存入数据库的 material_contents 表。

    :param file_list: Word文档的完整路径列表。
    :param material_id: 关联的教材ID。
    :param status_callback: 用于UI状态更新的回调函数。
    :param progress_callback: 用于UI进度条更新的回调函数。
    :return: 包含处理结果信息的字典。
    """
    result_summary = {
        "success": False, "message": "", "files_processed": 0, "total_files": len(file_list)
    }

    if not file_list:
        result_summary["message"] = "错误：文件列表为空。"
        return result_summary
        
    if not DB_CONFIG:
        result_summary["message"] = "错误：数据库配置 (db_config.py) 未找到或为空。"
        return result_summary

    word_app = None
    db_conn = None
    cursor = None
    coinitialized = False

    try:
        # --- 1. 初始化COM和Word ---
        pythoncom.CoInitialize()
        coinitialized = True
        status_callback("正在启动Word应用程序...")
        word_app = win32.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = False

        # --- 2. 连接数据库 ---
        status_callback("正在连接数据库...")
        db_conn = mysql.connector.connect(**DB_CONFIG)
        cursor = db_conn.cursor()
        status_callback("数据库连接成功。")

        # --- 3. 循环处理文件 ---
        status_callback(f"找到 {len(file_list)} 个待处理文档，准备开始处理...")

        for i, file_path in enumerate(file_list):
            progress_callback((i + 1) / len(file_list) * 100)
            filename = os.path.basename(file_path)
            status_callback(f"\n--- 开始处理文件: {filename} ({i+1}/{len(file_list)}) ---")
            
            # 为每个文件开启一个事务
            db_conn.start_transaction()

            doc = None
            try:
                doc = word_app.Documents.Open(file_path, ReadOnly=True)
                
                # 初始化状态变量，用于跟踪层级和内容
                last_inserted_id_at_level = {0: None} # key: level, value: DB id. Level 0 for root.
                sequence_counters = {None: 0}         # key: parent_id, value: next sequence number
                current_title_path = {}                # key: level, value: title text
                content_buffer = []                    # 收集正文段落
                last_heading_db_id = None              # 上一个标题在数据库中的ID，用于更新内容

                def flush_content_buffer():
                    """将缓冲区中的内容更新到上一个标题记录中"""
                    nonlocal last_heading_db_id
                    if content_buffer and last_heading_db_id:
                        full_content = "\n".join(content_buffer).strip()
                        if full_content:
                            status_callback(f"  -> 为ID {last_heading_db_id} 更新内容 (长度: {len(full_content)})...")
                            update_sql = "UPDATE material_contents SET content = %s, updated_at = %s WHERE id = %s"
                            cursor.execute(update_sql, (full_content, datetime.now(), last_heading_db_id))
                        content_buffer.clear()

                for para in doc.Paragraphs:
                    # 包含自动编号的完整段落文本
                    raw_text = para.Range.Text
                    list_string = para.Range.ListFormat.ListString
                    full_text = f"{list_string} {raw_text}" if list_string else raw_text
                    para_text = clean_text(full_text)

                    if not para_text:
                        continue
                    
                    level = para.OutlineLevel

                    # A. 如果是标题 (Level 1-9)
                    if 1 <= level <= 9:
                        # 先将之前收集的正文内容更新到上一个标题
                        flush_content_buffer()

                        # 准备插入新标题的数据
                        parent_level = level - 1
                        parent_id = last_inserted_id_at_level.get(parent_level)
                        
                        sequence = sequence_counters.get(parent_id, 0)
                        sequence_counters[parent_id] = sequence + 1
                        
                        current_title_path[level] = para_text
                        # 清理更深层级的旧路径
                        for l_key in list(current_title_path.keys()):
                            if l_key > level:
                                del current_title_path[l_key]
                        
                        # 准备冗余标题字段
                        titles = {f'title{l}': None for l in range(1, 9)}
                        for l in range(1, 9):
                            titles[f'title{l}'] = current_title_path.get(l)

                        # 插入数据库
                        insert_sql = """
                            INSERT INTO material_contents 
                            (material_id, parent_id, level, sequence, title, 
                             title1, title2, title3, title4, title5, title6, title7, title8,
                             content, summary, extra_data, created_at, updated_at)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, NULL, NULL, NULL, %s, %s)
                        """
                        now = datetime.now()
                        cursor.execute(insert_sql, (
                            material_id, parent_id, level, sequence, para_text,
                            titles['title1'], titles['title2'], titles['title3'], titles['title4'],
                            titles['title5'], titles['title6'], titles['title7'], titles['title8'],
                            now, now
                        ))
                        
                        new_id = cursor.lastrowid
                        last_inserted_id_at_level[level] = new_id
                        last_heading_db_id = new_id
                        status_callback(f"  -> 插入标题 (L{level}): '{para_text[:50]}...' -> ID: {new_id}, ParentID: {parent_id}")
                        
                        # 清理更深层级的ID记录
                        for l_key in list(last_inserted_id_at_level.keys()):
                            if l_key > level:
                                del last_inserted_id_at_level[l_key]

                    # B. 如果是正文
                    elif para_text: # 任何非标题的非空段落都视为内容
                        content_buffer.append(para_text)

                # 处理文档末尾的最后一部分内容
                flush_content_buffer()
                
                # 文件处理成功，提交事务
                db_conn.commit()
                status_callback(f"--- 文件 '{filename}' 处理完成，数据已提交。 ---")
                result_summary["files_processed"] += 1

            except Exception as e:
                status_callback(f"错误: 处理文件 '{filename}' 时发生严重错误: {e}")
                status_callback("正在回滚此文件的所有数据库更改...")
                if db_conn:
                    db_conn.rollback()
                status_callback("回滚完成。继续处理下一个文件。")
                continue # 继续处理下一个文件
            finally:
                if doc:
                    doc.Close(SaveChanges=False)

        result_summary["success"] = True
        result_summary["message"] = f"处理完成！共成功处理 {result_summary['files_processed']} / {result_summary['total_files']} 个文件。"

    except mysql.connector.Error as db_err:
        result_summary["message"] = f"数据库错误: {db_err}"
    except Exception as e:
        result_summary["message"] = f"发生未知错误: {e}"
        if db_conn:
            try: db_conn.rollback()
            except: pass
    finally:
        # --- 4. 清理资源 ---
        if cursor:
            cursor.close()
        if db_conn and db_conn.is_connected():
            db_conn.close()
            status_callback("数据库连接已关闭。")
        if word_app:
            word_app.Quit()
            status_callback("Word应用程序已关闭。")
        if coinitialized:
            pythoncom.CoUninitialize()

    return result_summary


# --- UI界面类 (已修改以适应数据库操作) ---
class WordExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word 内容提取并存入数据库 V6")
        self.root.geometry("700x600")

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
        
        # 新增: 教材ID输入框
        ttk.Label(self.action_frame, text="教材ID (material_id):").pack(side=tk.LEFT, padx=(10, 5))
        self.material_id_var = tk.StringVar()
        self.material_id_entry = ttk.Entry(self.action_frame, textvariable=self.material_id_var, width=10)
        self.material_id_entry.pack(side=tk.LEFT)
        
        self.start_button = ttk.Button(self.action_frame, text="开始提取并入库", command=self.start_extraction_thread)
        self.start_button.pack(side=tk.LEFT, padx=(10, 0))

        # 4. 状态和进度区
        self.status_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=12, state='disabled')
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

    def browse_directory(self):
        directory = filedialog.askdirectory(title="选择一个文件夹")
        if directory:
            current_files = self.file_listbox.get(0, tk.END)
            for filename in os.listdir(directory):
                if filename.endswith(('.doc', '.docx')) and not filename.startswith('~'):
                    full_path = os.path.join(directory, filename)
                    if full_path not in current_files:
                        self.file_listbox.insert(tk.END, full_path)

    def remove_selected_files(self):
        selected_indices = self.file_listbox.curselection()
        for i in reversed(selected_indices):
            self.file_listbox.delete(i)
            
    def clear_file_list(self):
        self.file_listbox.delete(0, tk.END)

    def log_status(self, message, is_error=False):
        self.root.after(0, self._log_status_sync, message, is_error)
        
    def _log_status_sync(self, message, is_error):
        self.status_text.config(state='normal')
        if is_error:
            self.status_text.insert(tk.END, message + "\n", 'error')
        else:
            self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')

    def update_progress(self, value):
        self.root.after(0, self.progress.config, {'value': value})
        
    def start_extraction_thread(self):
        file_list = self.file_listbox.get(0, tk.END)
        if not file_list:
            messagebox.showerror("错误", "请先选择至少一个文件或文件夹。")
            return

        material_id_str = self.material_id_var.get()
        if not material_id_str.isdigit():
            messagebox.showerror("错误", "请输入有效的数字作为教材ID (material_id)。")
            return
        material_id = int(material_id_str)
        
        if not DB_CONFIG:
            messagebox.showerror("配置错误", "数据库配置文件 'db_config.py' 未找到或为空，无法继续。")
            return

        self.start_button.config(state='disabled')
        self.progress['value'] = 0
        self.status_text.config(state='normal')
        self.status_text.delete('1.0', tk.END)
        self.status_text.config(state='disabled')

        thread = threading.Thread(
            target=self.run_extraction,
            args=(list(file_list), material_id)
        )
        thread.daemon = True
        thread.start()

    def run_extraction(self, file_list, material_id):
        result = extract_word_and_save_to_db(
            file_list, material_id,
            self.log_status,
            self.update_progress
        )
        self.root.after(0, self.on_extraction_complete, result)

    def on_extraction_complete(self, result):
        self.log_status("\n--- 任务总结 ---")
        is_error = not result['success']
        self.log_status(result['message'], is_error=is_error)

        self.start_button.config(state='normal')
        if result['success']:
             self.progress['value'] = 100
        else:
             self.progress['value'] = 0


if __name__ == '__main__':
    root = tk.Tk()
    app = WordExtractorApp(root)
    root.mainloop()