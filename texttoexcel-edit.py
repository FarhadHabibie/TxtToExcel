import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from ttkbootstrap.widgets import Button
from ttkbootstrap.constants import *
import os
import pandas as pd
from concurrent.futures import ThreadPoolExecutor


def read_tables_from_file(file_path):
    all_tables = []
    current_table = []
    is_table_started = False
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            line = line.strip()
            # stop before dispute section
            if ("LAPORAN TRANSAKSI DISPUTE INTERKONEKSI" in line 
                    or "LAPORAN TRANSAKSI DISPUTE" in line):
                break

            # toggle table on dashed lines
            if line.startswith('-') and set(line) == {'-'}:
                if is_table_started:
                    if current_table:
                        all_tables.append(current_table)
                    current_table = []
                    is_table_started = False
                else:
                    is_table_started = True
                continue

            if is_table_started and line:
                # split by tab if present; if no tabs, this will be a single-item list
                row = line.split('\t')
                # keep rows; don't assume the first row is header
                current_table.append(row)

    if current_table:
        all_tables.append(current_table)
    return all_tables


def read_dispute_tables_from_file(file_path):
    all_tables = []
    current_table = []
    is_table_started = False
    is_dispute_section = False

    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            line = line.strip()
            if ("LAPORAN TRANSAKSI DISPUTE INTERKONEKSI" in line 
                    or "LAPORAN TRANSAKSI DISPUTE" in line):
                is_dispute_section = True
                continue
            if not is_dispute_section:
                continue

            if line.startswith('-') and set(line) == {'-'}:
                if is_table_started:
                    if current_table:
                        all_tables.append(current_table)
                    current_table = []
                    is_table_started = False
                else:
                    is_table_started = True
                continue

            if is_table_started and line:
                current_table.append(line.split('\t'))

    if current_table:
        all_tables.append(current_table)

    # historical behavior: drop the very last trailing table if present
    return all_tables[:-1] if len(all_tables) > 1 else []


class FileUploaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Text File Transformer")
        self.root.geometry("600x450")
        self.style = Style("flatly")
        self.file_path = None
        self.executor = ThreadPoolExecutor(max_workers=2)

        Style("flatly")

        self.upload_btn = Button(root, text="Upload File", command=self.upload_file, bootstyle="primary")
        self.upload_btn.pack(pady=10)

        self.file_label = tk.Label(root, text="No file uploaded", fg="gray")
        self.file_label.pack(pady=5)

        self.remove_btn = Button(root, text="Remove File", command=self.remove_file, bootstyle="danger", state="disabled")
        self.remove_btn.pack(pady=5)

        self.run_btn1 = Button(root, text="Create Main Excel File", command=lambda: self.start_transformation(1), bootstyle="success", state="disabled")
        self.run_btn2 = Button(root, text="Create Dispute Excel File", command=lambda: self.start_transformation(2), bootstyle="success", state="disabled")
        self.run_btn1.pack(pady=10)
        self.run_btn2.pack(pady=10)

        self.status_label = tk.Label(root, text="", fg="blue", font=("Arial", 10, "bold"))
        self.status_label.pack(pady=5)

    def upload_file(self):
        path = filedialog.askopenfilename(filetypes=[("All Files", "*.*")])
        if path:
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    f.readline()
                self.file_path = path
                self.file_label.config(text=os.path.basename(path), fg="black")
                self.remove_btn.configure(state="normal")
                self.run_btn1.configure(state="normal")
                self.run_btn2.configure(state="normal")
            except Exception as e:
                messagebox.showerror("Unsupported File", f"Cannot read the selected file as text:\n{e}")

    def remove_file(self):
        self.file_path = None
        self.file_label.config(text="No file uploaded", fg="gray")
        self.remove_btn.configure(state="disabled")
        self.run_btn1.configure(state="disabled")
        self.run_btn2.configure(state="disabled")

    def start_transformation(self, mode):
        if not self.file_path:
            messagebox.showerror("Error", "No file uploaded.")
            return
        self.status_label.config(text="Processing...", fg="blue")
        if mode == 1:
            self.executor.submit(self.run_transformation)
        else:
            self.executor.submit(self.run_transformation2)

    @staticmethod
    def _build_df_from_tables_keep_all_rows(tables):
        """
        Combine all tables and make a single-column DataFrame 'col1'.
        If a row is already a list (from split), join it back to one string so the
        downstream fixed-width slicing works consistently.
        """
        combined = []
        for table in tables:
            combined.extend(table)

        # normalize rows to single string each
        rows_as_str = []
        for r in combined:
            if isinstance(r, list):
                rows_as_str.append('\t'.join(r))  # preserve tabs if they existed
            else:
                rows_as_str.append(str(r))

        # drop totally empty lines
        rows_as_str = [s for s in rows_as_str if s.strip() != ""]
        return pd.DataFrame(rows_as_str, columns=['col1'])

    @staticmethod
    def _strip_objects(df):
        for col in df.select_dtypes(include='object'):
            df[col] = df[col].str.strip()
        return df

    def run_transformation(self):
        try:
            tables = read_tables_from_file(self.file_path)

            # === KEEP ALL ROWS; DO NOT SLICE [1:] ===
            df = self._build_df_from_tables_keep_all_rows(tables)

            # Fixed-width parsing
            df['No.'] = df['col1'].str[:7]
            df['Trx_Code'] = df['col1'].str[7:16]
            df['Tanggal_Trx'] = df['col1'].str[16:28]
            df['Jam_Trx'] = df['col1'].str[28:37]
            df['Ref_No'] = df['col1'].str[37:50]
            df['Trace_No'] = df['col1'].str[50:59]
            df['Terminal_ID'] = df['col1'].str[59:76]
            df['Merchant_PAN'] = df['col1'].str[76:96]
            df['Acquirer'] = df['col1'].str[96:108]
            df['Issuer'] = df['col1'].str[108:120]
            df['Customer_PAN'] = df['col1'].str[120:140]
            df['Nominal'] = df['col1'].str[140:157]
            df['Merchant_Category'] = df['col1'].str[157:175]
            df['Merchant_Criteria'] = df['col1'].str[175:193]
            df['Response_Code'] = df['col1'].str[193:207]
            df['Merchant_Name'] = df['col1'].str[207:232]
            df['Location'] = df['col1'].str[232:245]
            df['Country'] = df['col1'].str[245:248]
            df['Convenience_Fee'] = df['col1'].str[248:261]
            df['Credit'] = df['col1'].str[261:263]
            df['Interchange_Fee'] = df['col1'].str[263:]

            # Clean objects
            df = self._strip_objects(df)

            # Keep only real data rows: 'No.' must be digits (prevents header lines)
            df = df[df['No.'].str.replace(' ', '').str.isdigit()]

            # Drop raw line
            df.drop(columns=['col1'], inplace=True)

            output_excel_path = os.path.splitext(self.file_path)[0] + '.xlsx'
            df.to_excel(output_excel_path, index=False)
            self.root.after(0, lambda: self.status_label.config(text="Done", fg="green"))
            self.root.after(0, lambda: messagebox.showinfo("Success", f"Excel file saved:\n{output_excel_path}"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Transformation failed:\n{e}"))

    def run_transformation2(self):
        try:
            tables = read_dispute_tables_from_file(self.file_path)
            if not tables:
                raise IndexError

            # === KEEP ALL ROWS; DO NOT SLICE [1:] ===
            df = self._build_df_from_tables_keep_all_rows(tables)

            # Fixed-width parsing (dispute layout)
            df['No.'] = df['col1'].str[:7]
            df['Trx_Code'] = df['col1'].str[7:16]
            df['Tanggal_Trx'] = df['col1'].str[16:28]
            df['Jam_Trx'] = df['col1'].str[28:37]
            df['Ref_No'] = df['col1'].str[37:50]
            df['Trace_No'] = df['col1'].str[50:59]
            df['Terminal_ID'] = df['col1'].str[59:76]
            df['Merchant_PAN'] = df['col1'].str[76:96]
            df['Acquirer'] = df['col1'].str[96:108]
            df['Issuer'] = df['col1'].str[108:120]
            df['Customer_PAN'] = df['col1'].str[120:140]
            df['Nominal'] = df['col1'].str[140:157]
            df['Merchant_Category'] = df['col1'].str[157:175]
            df['Merchant_Criteria'] = df['col1'].str[175:193]
            df['Response_Code'] = df['col1'].str[193:207]
            df['Merchant_Name'] = df['col1'].str[207:232]
            df['Location'] = df['col1'].str[232:245]
            df['Country'] = df['col1'].str[245:248]
            df['Convenience_Fee'] = df['col1'].str[248:261]
            df['Credit'] = df['col1'].str[261:263]
            df['Interchange_Fee'] = df['col1'].str[263:279]
            df['Dispute_Tran_Code'] = df['col1'].str[279:299]
            df['Dispute_Amount'] = df['col1'].str[299:317]
            df['Fee_Return'] = df['col1'].str[317:335]
            df['Dispute_Net_Amount'] = df['col1'].str[335:355]
            df['Registration_Number'] = df['col1'].str[355:374]

            # Clean objects
            df = self._strip_objects(df)

            # Keep only real data rows: 'No.' must be digits (prevents header lines)
            df = df[df['No.'].str.replace(' ', '').str.isdigit()]

            # Drop raw line
            df.drop(columns=['col1'], inplace=True)

            output_excel_path = os.path.join(
                os.path.dirname(self.file_path),
                'DISPUTE_' + os.path.splitext(os.path.basename(self.file_path))[0] + '.xlsx'
            )
            df.to_excel(output_excel_path, index=False)
            self.root.after(0, lambda: self.status_label.config(text="Done", fg="green"))
            self.root.after(0, lambda: messagebox.showinfo("Success", f"Excel file saved:\n{output_excel_path}"))
        except IndexError:
            self.root.after(0, lambda: messagebox.showerror("Error", "No data found in the dispute section."))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Transformation 2 failed:\n{e}"))


if __name__ == "__main__":
    root = tk.Tk()
    app = FileUploaderApp(root)
    root.mainloop()
