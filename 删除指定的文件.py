import tkinter as tk
from tkinter import scrolledtext, filedialog, messagebox
import os

class FileDeleterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("文件批量删除工具")
        self.geometry("600x500")
        self.resizable(False, False) # 不允许调整窗口大小

        # 1. 说明标签
        self.label = tk.Label(self, text="请输入要删除的文件路径（每行一个，支持粘贴）：", font=("Helvetica", 10))
        self.label.pack(pady=10)

        # 2. 文件列表输入框 (带滚动条的文本框)
        self.file_list_text = scrolledtext.ScrolledText(self, width=70, height=15, wrap=tk.WORD, font=("Consolas", 9))
        self.file_list_text.pack(pady=5)
        self.file_list_text.focus_set() # 启动时让输入框获得焦点

        # 3. 按钮框架
        self.button_frame = tk.Frame(self)
        self.button_frame.pack(pady=10)

        self.browse_button = tk.Button(self.button_frame, text="浏览文件...", command=self.browse_files, font=("Helvetica", 9))
        self.browse_button.pack(side=tk.LEFT, padx=10)

        self.clear_button = tk.Button(self.button_frame, text="清空", command=self.clear_input, font=("Helvetica", 9))
        self.clear_button.pack(side=tk.LEFT, padx=10)

        self.delete_button = tk.Button(self.button_frame, text="确认删除", command=self.confirm_deletion, 
                                       bg="red", fg="white", font=("Helvetica", 10, "bold"),
                                       activebackground="#CC0000", activeforeground="white")
        self.delete_button.pack(side=tk.LEFT, padx=20)

        # 4. 状态信息标签
        self.status_label = tk.Label(self, text="等待操作...", fg="blue", font=("Helvetica", 10))
        self.status_label.pack(pady=10)

    def browse_files(self):
        """打开文件选择对话框，将选中的文件路径添加到文本框"""
        filepaths = filedialog.askopenfilenames(
            title="选择要删除的文件",
            filetypes=[("所有文件", "*.*"), ("文本文件", "*.txt"), ("图像文件", "*.png;*.jpg;*.jpeg")]
        )
        if filepaths:
            for fp in filepaths:
                # 检查是否已存在，避免重复添加（简单检查）
                current_text = self.file_list_text.get(1.0, tk.END)
                if fp + "\n" not in current_text and fp not in current_text.splitlines():
                    self.file_list_text.insert(tk.END, fp + "\n")
            self.status_label.config(text=f"已添加 {len(filepaths)} 个文件路径。")

    def clear_input(self):
        """清空文件列表输入框的内容"""
        self.file_list_text.delete(1.0, tk.END) # 1.0表示第一行第一个字符到末尾
        self.status_label.config(text="输入框已清空。")

    def confirm_deletion(self):
        """获取文件列表，弹出确认对话框，然后执行删除操作"""
        # 获取文本框内容，去除首尾空白，并按行分割
        file_paths_raw = self.file_list_text.get(1.0, tk.END).strip()

        if not file_paths_raw:
            messagebox.showwarning("警告", "文件列表为空，请输入要删除的文件路径。")
            self.status_label.config(text="文件列表为空。", fg="orange")
            return

        # 清理路径：去除空行和每行首尾空白
        file_paths = [p.strip() for p in file_paths_raw.split('\n') if p.strip()]

        if not file_paths:
            messagebox.showwarning("警告", "无效的文件路径，请检查输入。")
            self.status_label.config(text="无效的文件路径。", fg="orange")
            return

        # 构建确认消息
        msg_files_preview = "\n".join(file_paths[:min(5, len(file_paths))]) # 最多显示前5个文件
        if len(file_paths) > 5:
            msg_files_preview += "\n..."
        
        confirmation_msg = (
            f"您确定要删除以下 {len(file_paths)} 个文件吗？\n\n"
            f"{msg_files_preview}\n\n"
            "此操作不可逆！请谨慎确认！"
        )

        # 弹出确认对话框
        if not messagebox.askyesno("确认删除", confirmation_msg):
            self.status_label.config(text="删除操作已取消。", fg="red")
            return

        # 执行删除
        self.status_label.config(text="正在删除文件，请稍候...", fg="blue")
        self.update_idletasks() # 强制更新UI，让用户看到“正在删除”

        deleted_count = 0
        failed_deletions = []

        for path in file_paths:
            if not path: # 再次确认路径非空
                continue
            
            try:
                if os.path.exists(path):
                    os.remove(path)
                    deleted_count += 1
                else:
                    failed_deletions.append(f"'{path}' (文件或目录不存在)")
            except FileNotFoundError: # 理论上被 os.path.exists 捕获，但作为备用
                failed_deletions.append(f"'{path}' (未找到文件)")
            except PermissionError:
                failed_deletions.append(f"'{path}' (权限不足，无法删除)")
            except OSError as e: # 其他OS错误，如目录不是空的，或者路径太长等
                failed_deletions.append(f"'{path}' (删除失败: {e})")

        # 结果反馈
        if not failed_deletions:
            final_status_msg = f"成功删除 {deleted_count} 个文件！"
            self.status_label.config(text=final_status_msg, fg="green")
            messagebox.showinfo("删除完成", final_status_msg)
            self.clear_input() # 成功删除后清空输入框
        else:
            final_status_msg = f"删除完成。成功删除 {deleted_count} 个文件。\n\n"
            final_status_msg += "以下文件删除失败：\n" + "\n".join(failed_deletions)
            self.status_label.config(text=f"部分文件删除失败，请查看详情。", fg="red")
            messagebox.showerror("删除失败", final_status_msg)

if __name__ == "__main__":
    app = FileDeleterApp()
    app.mainloop()