import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
from datetime import datetime

class FileInfoExporter:
    """
    一个用于导出目录文件信息到Excel的GUI应用
    """
    def __init__(self, master):
        self.master = master
        master.title("文件信息导出工具")
        master.geometry("500x250") # 设置窗口初始大小

        # --- GUI 控件 ---

        # 1. 源目录选择
        self.source_dir_label = tk.Label(master, text="选择源目录:")
        self.source_dir_label.grid(row=0, column=0, padx=10, pady=10, sticky='w')

        self.source_dir_var = tk.StringVar()
        self.source_dir_entry = tk.Entry(master, textvariable=self.source_dir_var, width=50)
        self.source_dir_entry.grid(row=0, column=1, padx=10, pady=10)

        self.browse_source_btn = tk.Button(master, text="浏览...", command=self.browse_source_dir)
        self.browse_source_btn.grid(row=0, column=2, padx=10, pady=10)

        # 2. 导出文件选择
        self.export_file_label = tk.Label(master, text="导出到文件:")
        self.export_file_label.grid(row=1, column=0, padx=10, pady=10, sticky='w')

        self.export_file_var = tk.StringVar()
        self.export_file_entry = tk.Entry(master, textvariable=self.export_file_var, width=50)
        self.export_file_entry.grid(row=1, column=1, padx=10, pady=10)

        self.browse_export_btn = tk.Button(master, text="另存为...", command=self.browse_export_file)
        self.browse_export_btn.grid(row=1, column=2, padx=10, pady=10)

        # 3. 执行按钮
        self.export_btn = tk.Button(master, text="开始导出", command=self.export_to_excel, font=("Arial", 12, "bold"), bg="lightblue")
        self.export_btn.grid(row=2, column=1, pady=20)

        # 4. 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("准备就绪")
        self.status_label = tk.Label(master, textvariable=self.status_var, fg="blue", anchor='w')
        self.status_label.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky='ew')
        
    def browse_source_dir(self):
        """打开对话框选择源目录"""
        directory = filedialog.askdirectory()
        if directory:
            self.source_dir_var.set(directory)
            self.status_var.set(f"已选择源目录: {directory}")

    def browse_export_file(self):
        """打开“另存为”对话框选择导出文件路径"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
            title="保存到 Excel 文件"
        )
        if filename:
            self.export_file_var.set(filename)
            self.status_var.set(f"将导出到: {filename}")

    def get_file_info(self, directory):
        """获取目录下所有文件的信息"""
        file_data = []
        for filename in os.listdir(directory):
            full_path = os.path.join(directory, filename)
            
            # 确保只处理文件，忽略子目录
            if os.path.isfile(full_path):
                try:
                    stats = os.stat(full_path)
                    
                    # 获取文件大小并转换为KB
                    size_kb = round(stats.st_size / 1024, 2)
                    
                    # 获取文件修改日期
                    mod_time = datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                    
                    # 获取文件名和扩展名
                    _, file_extension = os.path.splitext(filename)
                    
                    file_info = {
                        '目录名': directory,
                        '文件名': filename,
                        '文件大小(KB)': size_kb,
                        '文件类型': file_extension,
                        '修改日期': mod_time
                    }
                    file_data.append(file_info)
                except Exception as e:
                    print(f"无法访问文件 {full_path}: {e}") # 在控制台打印错误，避免中断
        return file_data

    def export_to_excel(self):
        """执行导出操作"""
        source_dir = self.source_dir_var.get()
        export_path = self.export_file_var.get()

        # 1. 验证输入
        if not source_dir or not os.path.isdir(source_dir):
            messagebox.showerror("错误", "请选择一个有效的源目录！")
            return
        
        if not export_path:
            messagebox.showerror("错误", "请指定要导出的文件名！")
            return

        try:
            self.status_var.set("正在处理，请稍候...")
            self.master.update_idletasks() # 强制更新UI

            # 2. 获取文件信息
            file_list = self.get_file_info(source_dir)

            if not file_list:
                messagebox.showwarning("提示", "所选目录中没有找到任何文件。")
                self.status_var.set("准备就绪")
                return

            # 3. 创建DataFrame并导出到Excel
            df = pd.DataFrame(file_list)
            
            # 重新排列列的顺序，以符合要求
            df = df[['目录名', '文件名', '文件大小(KB)', '文件类型', '修改日期']]
            
            # 写入Excel文件，不包含索引列
            df.to_excel(export_path, index=False, engine='openpyxl')
            
            self.status_var.set(f"成功！已导出 {len(file_list)} 个文件信息。")
            messagebox.showinfo("成功", f"文件信息已成功导出到:\n{export_path}")

        except Exception as e:
            self.status_var.set(f"导出失败: {e}")
            messagebox.showerror("导出失败", f"发生错误:\n{e}")

# --- 主程序入口 ---
if __name__ == "__main__":
    root = tk.Tk()
    app = FileInfoExporter(root)
    root.mainloop()