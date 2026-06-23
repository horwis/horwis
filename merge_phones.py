import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import traceback
from datetime import datetime


class MergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("手机号合并去重工具")
        self.root.geometry("620x320")

        self.high_file = tk.StringVar()
        self.mid_file = tk.StringVar()
        self.low_file = tk.StringVar()
        self.out_file = tk.StringVar()
        self.output_format = tk.StringVar(value="Excel (.xlsx)")  # 默认 Excel

        self.create_widgets()

    def create_widgets(self):
        # 高优先级
        tk.Label(self.root, text="高优先级文件 (可选):").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        tk.Entry(self.root, textvariable=self.high_file, width=50).grid(row=0, column=1, padx=5)
        tk.Button(self.root, text="浏览", command=lambda: self.browse_file(self.high_file)).grid(row=0, column=2,
                                                                                                 padx=5)

        # 中优先级
        tk.Label(self.root, text="中优先级文件 (可选):").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        tk.Entry(self.root, textvariable=self.mid_file, width=50).grid(row=1, column=1, padx=5)
        tk.Button(self.root, text="浏览", command=lambda: self.browse_file(self.mid_file)).grid(row=1, column=2, padx=5)

        # 低优先级
        tk.Label(self.root, text="低优先级文件 (可选):").grid(row=2, column=0, sticky='e', padx=5, pady=5)
        tk.Entry(self.root, textvariable=self.low_file, width=50).grid(row=2, column=1, padx=5)
        tk.Button(self.root, text="浏览", command=lambda: self.browse_file(self.low_file)).grid(row=2, column=2, padx=5)

        # 输出文件
        tk.Label(self.root, text="输出文件路径 (必选):").grid(row=3, column=0, sticky='e', padx=5, pady=5)
        tk.Entry(self.root, textvariable=self.out_file, width=50).grid(row=3, column=1, padx=5)
        tk.Button(self.root, text="浏览", command=self.browse_save_file).grid(row=3, column=2, padx=5)

        # 输出格式选择
        tk.Label(self.root, text="输出格式:").grid(row=4, column=0, sticky='e', padx=5, pady=5)
        format_combo = ttk.Combobox(self.root, textvariable=self.output_format,
                                    values=["Excel (.xlsx)", "CSV (.csv)"], state="readonly")
        format_combo.grid(row=4, column=1, sticky='w', padx=5)
        format_combo.bind("<<ComboboxSelected>>", self.on_format_change)

        # 提示信息
        tk.Label(self.root, text="注意：至少选择一个输入文件，输出必须指定。", fg="blue").grid(row=5, column=0,
                                                                                            columnspan=3, pady=5)

        # 处理按钮
        self.process_btn = tk.Button(self.root, text="开始处理", command=self.start_processing, width=15)
        self.process_btn.grid(row=6, column=1, pady=10)

        # 状态与进度条
        self.status_label = tk.Label(self.root, text="就绪")
        self.status_label.grid(row=7, column=0, columnspan=3, pady=5)
        self.progress = ttk.Progressbar(self.root, orient='horizontal', length=400, mode='indeterminate')
        self.progress.grid(row=8, column=0, columnspan=3, pady=5)

    def browse_file(self, var):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            var.set(filename)

    def browse_save_file(self):
        # 根据当前选择的格式决定默认扩展名和文件类型
        if self.output_format.get() == "Excel (.xlsx)":
            def_ext = ".xlsx"
            filetypes = [("Excel files", "*.xlsx"), ("All files", "*.*")]
        else:  # CSV
            def_ext = ".csv"
            filetypes = [("CSV files", "*.csv"), ("All files", "*.*")]

        filename = filedialog.asksaveasfilename(
            defaultextension=def_ext,
            filetypes=filetypes
        )
        if filename:
            # 如果用户未输入扩展名，自动添加
            if def_ext == ".xlsx" and not filename.lower().endswith(".xlsx"):
                filename += ".xlsx"
            elif def_ext == ".csv" and not filename.lower().endswith(".csv"):
                filename += ".csv"
            self.out_file.set(filename)

    def on_format_change(self, event=None):
        # 当切换格式时，如果当前输出路径已有值，可自动更改后缀（可选）
        current = self.out_file.get().strip()
        if current:
            # 简单处理：若当前路径有扩展名则替换，否则不处理
            base, ext = os.path.splitext(current)
            if self.output_format.get() == "Excel (.xlsx)":
                new_path = base + ".xlsx"
            else:
                new_path = base + ".csv"
            if new_path != current:
                self.out_file.set(new_path)

    def start_processing(self):
        high = self.high_file.get().strip()
        mid = self.mid_file.get().strip()
        low = self.low_file.get().strip()
        out = self.out_file.get().strip()

        if not (high or mid or low):
            messagebox.showerror("错误", "至少选择一个输入文件！")
            return
        if not out:
            messagebox.showerror("错误", "请指定输出文件路径！")
            return

        for fpath in [high, mid, low]:
            if fpath and not os.path.exists(fpath):
                messagebox.showerror("错误", f"文件不存在：{fpath}")
                return

        self.process_btn.config(state='disabled')
        self.status_label.config(text="处理中...")
        self.progress.start()
        thread = threading.Thread(target=self.process_files, args=(high, mid, low, out))
        thread.daemon = True
        thread.start()

    def write_excel_with_split(self, df, filepath, max_rows=1048576):
        """将 DataFrame 写入 Excel，自动拆分为多个 sheet，返回 sheet 数量"""
        if len(df) == 0:
            df.to_excel(filepath, index=False, engine='openpyxl')
            return 1
        if len(df) <= max_rows:
            df.to_excel(filepath, index=False, engine='openpyxl')
            return 1
        else:
            num_sheets = (len(df) + max_rows - 1) // max_rows
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                for i in range(num_sheets):
                    start = i * max_rows
                    end = min((i + 1) * max_rows, len(df))
                    sheet_name = f"Sheet{i + 1}"
                    df.iloc[start:end].to_excel(writer, sheet_name=sheet_name, index=False)
            return num_sheets

    def process_files(self, high, mid, low, out):
        try:
            def read_excel_file(filepath, priority):
                if not filepath:
                    return pd.DataFrame(columns=['手机号', '优先级'])
                sheets = pd.read_excel(filepath, sheet_name=None, header=None, dtype=str)
                df_list = []
                for sheet_name, df in sheets.items():
                    if df.shape[1] >= 1:
                        col = df.iloc[:, 0]
                        col = col.dropna()
                        col = col.astype(str).str.strip()
                        col = col[col != '']
                        temp_df = pd.DataFrame({'手机号': col, '优先级': priority})
                        df_list.append(temp_df)
                if df_list:
                    return pd.concat(df_list, ignore_index=True)
                else:
                    return pd.DataFrame(columns=['手机号', '优先级'])

            def update_status(msg):
                self.root.after(0, lambda: self.status_label.config(text=msg))

            selected = []
            if high:
                selected.append(('高', high))
            if mid:
                selected.append(('中', mid))
            if low:
                selected.append(('低', low))

            df_list = []
            for priority, fpath in selected:
                update_status(f"读取 {priority} 优先级文件...")
                df = read_excel_file(fpath, priority)
                if not df.empty:
                    df_list.append(df)

            if not df_list:
                raise Exception("所有文件均为空或无可读数据！")

            update_status("合并数据...")
            df_all = pd.concat(df_list, ignore_index=True)
            df_all.drop_duplicates(subset=['手机号', '优先级'], keep='first', inplace=True)

            unique_priorities = df_all['优先级'].unique()
            if len(unique_priorities) == 1:
                df_all.drop_duplicates(subset=['手机号'], keep='first', inplace=True)
            else:
                priority_map = {'高': 3, '中': 2, '低': 1}
                df_all['等级'] = df_all['优先级'].map(priority_map)
                idx = df_all.groupby('手机号')['等级'].idxmax()
                df_all = df_all.loc[idx]
                df_all.drop('等级', axis=1, inplace=True)

            df_all.reset_index(drop=True, inplace=True)

            # 根据输出格式写入
            update_status("写入文件...")
            fmt = self.output_format.get()
            if fmt == "Excel (.xlsx)":
                sheet_count = self.write_excel_with_split(df_all, out)
                extra_msg = f"已拆分为 {sheet_count} 个工作表。" if sheet_count > 1 else ""
            else:  # CSV
                # 写入 CSV，使用 utf-8-sig 编码以便 Excel 打开无乱码
                df_all.to_csv(out, index=False, encoding='utf-8-sig')
                extra_msg = "CSV 格式无行数限制。"
                sheet_count = 1

            self.root.after(0, lambda: self.status_label.config(text="处理完成！"))
            self.root.after(0, lambda: messagebox.showinfo(
                "完成",
                f"处理完成！\n输出文件：{out}\n共保留 {len(df_all)} 条记录。\n{extra_msg}"
            ))

        except Exception as e:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_filename = f"error_{timestamp}.log"
            tb_text = traceback.format_exc()
            try:
                with open(log_filename, 'w', encoding='utf-8') as f:
                    f.write(f"错误发生时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"错误信息：{str(e)}\n")
                    f.write("\n详细堆栈跟踪：\n")
                    f.write(tb_text)
                log_msg = f"错误已记录到文件：{os.path.abspath(log_filename)}"
            except Exception as log_err:
                log_msg = f"无法写入日志文件：{log_err}"

            self.root.after(0, lambda: self.status_label.config(text="出错"))
            self.root.after(0, lambda: messagebox.showerror("错误", f"处理出错！\n{str(e)}\n\n{log_msg}"))
        finally:
            self.root.after(0, lambda: self.process_btn.config(state='normal'))
            self.root.after(0, lambda: self.progress.stop())


if __name__ == "__main__":
    root = tk.Tk()
    app = MergeApp(root)
    root.mainloop()