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
        self.root.geometry("700x500")

        # 存储每个文件的路径、表头选项、列号
        self.high_file = tk.StringVar()
        self.mid_file = tk.StringVar()
        self.low_file = tk.StringVar()
        self.high_header = tk.StringVar(value="是")
        self.mid_header = tk.StringVar(value="是")
        self.low_header = tk.StringVar(value="是")
        self.high_col = tk.StringVar(value="1")
        self.mid_col = tk.StringVar(value="1")
        self.low_col = tk.StringVar(value="1")

        self.out_file = tk.StringVar()
        self.output_format = tk.StringVar(value="Excel (.xlsx)")

        self.create_widgets()

    def create_widgets(self):
        # 高优先级区域
        frame_high = tk.LabelFrame(self.root, text="高优先级文件", padx=5, pady=5)
        frame_high.grid(row=0, column=0, columnspan=3, sticky='ew', padx=5, pady=5)
        tk.Label(frame_high, text="文件:").grid(row=0, column=0, sticky='e')
        tk.Entry(frame_high, textvariable=self.high_file, width=50).grid(row=0, column=1, padx=5)
        tk.Button(frame_high, text="浏览", command=lambda: self.browse_file(self.high_file)).grid(row=0, column=2)
        # 选项行
        tk.Label(frame_high, text="是否包含表头:").grid(row=1, column=0, sticky='e')
        header_combo = ttk.Combobox(frame_high, textvariable=self.high_header, values=["是", "否"], state="readonly",
                                    width=5)
        header_combo.grid(row=1, column=1, sticky='w', padx=5)
        tk.Label(frame_high, text="手机号所在列 (序号):").grid(row=1, column=2, sticky='e', padx=(10, 0))
        tk.Entry(frame_high, textvariable=self.high_col, width=6).grid(row=1, column=3, sticky='w')

        # 中优先级区域
        frame_mid = tk.LabelFrame(self.root, text="中优先级文件", padx=5, pady=5)
        frame_mid.grid(row=1, column=0, columnspan=3, sticky='ew', padx=5, pady=5)
        tk.Label(frame_mid, text="文件:").grid(row=0, column=0, sticky='e')
        tk.Entry(frame_mid, textvariable=self.mid_file, width=50).grid(row=0, column=1, padx=5)
        tk.Button(frame_mid, text="浏览", command=lambda: self.browse_file(self.mid_file)).grid(row=0, column=2)
        tk.Label(frame_mid, text="是否包含表头:").grid(row=1, column=0, sticky='e')
        header_combo = ttk.Combobox(frame_mid, textvariable=self.mid_header, values=["是", "否"], state="readonly",
                                    width=5)
        header_combo.grid(row=1, column=1, sticky='w', padx=5)
        tk.Label(frame_mid, text="手机号所在列 (序号):").grid(row=1, column=2, sticky='e', padx=(10, 0))
        tk.Entry(frame_mid, textvariable=self.mid_col, width=6).grid(row=1, column=3, sticky='w')

        # 低优先级区域
        frame_low = tk.LabelFrame(self.root, text="低优先级文件", padx=5, pady=5)
        frame_low.grid(row=2, column=0, columnspan=3, sticky='ew', padx=5, pady=5)
        tk.Label(frame_low, text="文件:").grid(row=0, column=0, sticky='e')
        tk.Entry(frame_low, textvariable=self.low_file, width=50).grid(row=0, column=1, padx=5)
        tk.Button(frame_low, text="浏览", command=lambda: self.browse_file(self.low_file)).grid(row=0, column=2)
        tk.Label(frame_low, text="是否包含表头:").grid(row=1, column=0, sticky='e')
        header_combo = ttk.Combobox(frame_low, textvariable=self.low_header, values=["是", "否"], state="readonly",
                                    width=5)
        header_combo.grid(row=1, column=1, sticky='w', padx=5)
        tk.Label(frame_low, text="手机号所在列 (序号):").grid(row=1, column=2, sticky='e', padx=(10, 0))
        tk.Entry(frame_low, textvariable=self.low_col, width=6).grid(row=1, column=3, sticky='w')

        # 输出文件与格式
        frame_out = tk.LabelFrame(self.root, text="输出设置", padx=5, pady=5)
        frame_out.grid(row=3, column=0, columnspan=3, sticky='ew', padx=5, pady=5)
        tk.Label(frame_out, text="输出文件:").grid(row=0, column=0, sticky='e')
        tk.Entry(frame_out, textvariable=self.out_file, width=50).grid(row=0, column=1, padx=5)
        tk.Button(frame_out, text="浏览", command=self.browse_save_file).grid(row=0, column=2)
        tk.Label(frame_out, text="输出格式:").grid(row=1, column=0, sticky='e')
        format_combo = ttk.Combobox(frame_out, textvariable=self.output_format,
                                    values=["Excel (.xlsx)", "CSV (.csv)"], state="readonly")
        format_combo.grid(row=1, column=1, sticky='w', padx=5)
        format_combo.bind("<<ComboboxSelected>>", self.on_format_change)

        # 提示与按钮
        tk.Label(self.root, text="注意：至少选择一个输入文件，输出必须指定。", fg="blue").grid(row=4, column=0,
                                                                                            columnspan=3, pady=5)
        self.process_btn = tk.Button(self.root, text="开始处理", command=self.start_processing, width=15)
        self.process_btn.grid(row=5, column=1, pady=10)

        self.status_label = tk.Label(self.root, text="就绪")
        self.status_label.grid(row=6, column=0, columnspan=3, pady=5)
        self.progress = ttk.Progressbar(self.root, orient='horizontal', length=500, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=3, pady=5)

    def browse_file(self, var):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            var.set(filename)

    def browse_save_file(self):
        if self.output_format.get() == "Excel (.xlsx)":
            def_ext = ".xlsx"
            filetypes = [("Excel files", "*.xlsx"), ("All files", "*.*")]
        else:
            def_ext = ".csv"
            filetypes = [("CSV files", "*.csv"), ("All files", "*.*")]
        filename = filedialog.asksaveasfilename(defaultextension=def_ext, filetypes=filetypes)
        if filename:
            if def_ext == ".xlsx" and not filename.lower().endswith(".xlsx"):
                filename += ".xlsx"
            elif def_ext == ".csv" and not filename.lower().endswith(".csv"):
                filename += ".csv"
            self.out_file.set(filename)

    def on_format_change(self, event=None):
        current = self.out_file.get().strip()
        if current:
            base, ext = os.path.splitext(current)
            if self.output_format.get() == "Excel (.xlsx)":
                new_path = base + ".xlsx"
            else:
                new_path = base + ".csv"
            if new_path != current:
                self.out_file.set(new_path)

    def start_processing(self):
        # 收集文件和参数
        high = self.high_file.get().strip()
        mid = self.mid_file.get().strip()
        low = self.low_file.get().strip()
        out = self.out_file.get().strip()

        # 至少一个输入
        if not (high or mid or low):
            messagebox.showerror("错误", "至少选择一个输入文件！")
            return
        if not out:
            messagebox.showerror("错误", "请指定输出文件路径！")
            return

        # 校验列号
        try:
            high_col = int(self.high_col.get().strip()) if high else 1
            mid_col = int(self.mid_col.get().strip()) if mid else 1
            low_col = int(self.low_col.get().strip()) if low else 1
            if high_col < 1 or mid_col < 1 or low_col < 1:
                raise ValueError("列号必须为正整数")
        except ValueError as e:
            messagebox.showerror("错误", f"列号输入无效：{str(e)}")
            return

        # 保存参数供后续使用
        self.file_params = []
        if high:
            self.file_params.append(('高', high, self.high_header.get() == "是", high_col))
        if mid:
            self.file_params.append(('中', mid, self.mid_header.get() == "是", mid_col))
        if low:
            self.file_params.append(('低', low, self.low_header.get() == "是", low_col))

        self.process_btn.config(state='disabled')
        self.status_label.config(text="处理中...")
        self.progress.start()
        thread = threading.Thread(target=self.process_files, args=(out,))
        thread.daemon = True
        thread.start()

    def write_excel_with_split(self, df, filepath, max_rows=1048576):
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

    def process_files(self, out):
        try:
            def read_excel_file(filepath, priority, has_header, col_idx):
                """
                读取 Excel 文件所有 sheets，提取指定列（col_idx 为 1-based）的数据。
                has_header: True 表示第一行为表头，数据从第二行开始；False 表示所有行都是数据。
                返回 DataFrame 包含 '手机号' 和 '优先级'
                """
                sheets = pd.read_excel(filepath, sheet_name=None, header=None, dtype=str)
                df_list = []
                for sheet_name, df in sheets.items():
                    if df.shape[1] < col_idx:
                        raise Exception(f"工作表 '{sheet_name}' 列数不足，指定列 {col_idx} 不存在（总列数 {df.shape[1]}）")
                    # 提取指定列（注意 col_idx 是 1-based，转为 0-based）
                    col_data = df.iloc[:, col_idx - 1]
                    # 根据是否包含表头决定起始行
                    start_row = 1 if has_header else 0
                    # 截取数据（从 start_row 开始）
                    if len(col_data) <= start_row:
                        continue  # 无数据
                    col_data = col_data.iloc[start_row:].dropna()
                    col_data = col_data.astype(str).str.strip()
                    col_data = col_data[col_data != '']
                    if len(col_data) > 0:
                        temp_df = pd.DataFrame({'手机号': col_data.values, '优先级': priority})
                        df_list.append(temp_df)
                if df_list:
                    return pd.concat(df_list, ignore_index=True)
                else:
                    return pd.DataFrame(columns=['手机号', '优先级'])

            def update_status(msg):
                self.root.after(0, lambda: self.status_label.config(text=msg))

            df_list = []
            for priority, fpath, has_header, col_idx in self.file_params:
                update_status(f"读取 {priority} 优先级文件...")
                df = read_excel_file(fpath, priority, has_header, col_idx)
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

            # 写入
            update_status("写入文件...")
            fmt = self.output_format.get()
            if fmt == "Excel (.xlsx)":
                sheet_count = self.write_excel_with_split(df_all, out)
                extra_msg = f"已拆分为 {sheet_count} 个工作表。" if sheet_count > 1 else ""
            else:
                df_all.to_csv(out, index=False, encoding='utf-8-sig')
                extra_msg = "CSV 格式无行数限制。"

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
