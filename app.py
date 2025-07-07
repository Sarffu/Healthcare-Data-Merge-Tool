import pandas as pd
from tkinter import Tk, filedialog, ttk, messagebox
import tkinter as tk
import openpyxl
from openpyxl.styles import PatternFill
import threading
import os

# Global variables
file1_path_global = None
file2_path_global = None
save_path_global = None
merged_df_global = None

def select_file1():
    global file1_path_global
    file1_path_global = filedialog.askopenfilename(
        title="Select Schedule File (with NPI & VotedDate)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file1_path_global:
        file1_status.config(text=f"Selected: {os.path.basename(file1_path_global)}", foreground="green")
    else:
        file1_status.config(text="No file selected", foreground="red")
    check_and_enable_merge_button()

def select_file2():
    global file2_path_global
    file2_path_global = filedialog.askopenfilename(
        title="Select Roaster File (with Individual NPI & Provider Effective Date)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file2_path_global:
        file2_status.config(text=f"Selected: {os.path.basename(file2_path_global)}", foreground="green")
    else:
        file2_status.config(text="No file selected", foreground="red")
    check_and_enable_merge_button()

def check_and_enable_merge_button():
    if file1_path_global and file2_path_global:
        merge_button.config(state=tk.NORMAL, background="#4CAF50")
        status_label.config(text="Ready to merge files", foreground="green")
    else:
        merge_button.config(state=tk.DISABLED, background="lightgray")
        status_label.config(text="Please select both files to continue", foreground="orange")

def perform_merge_logic():
    global save_path_global, merged_df_global

    window.after(0, lambda: status_label.config(text="Processing merge...", foreground="blue"))
    window.after(0, lambda: merge_button.config(state=tk.DISABLED))
    window.after(0, lambda: select_file1_button.config(state=tk.DISABLED))
    window.after(0, lambda: select_file2_button.config(state=tk.DISABLED))

    try:
        df1 = pd.read_excel(file1_path_global)
        df2 = pd.read_excel(file2_path_global)

        df1.columns = df1.columns.str.strip()
        df2.columns = df2.columns.str.strip()

        if 'NPI' not in df1.columns or 'VotedDate' not in df1.columns:
            raise ValueError("Schedule file must contain 'NPI' and 'VotedDate' columns.")
        if 'Individual NPI' not in df2.columns or 'Provider Effective Date' not in df2.columns:
            raise ValueError("Roaster file must contain 'Individual NPI' and 'Provider Effective Date' columns.")

        df2_original_dates = df2[['Individual NPI', 'Provider Effective Date']].copy()
        merged = pd.merge(df2, df1[['NPI', 'VotedDate']], how='left', left_on='Individual NPI', right_on='NPI')
        merged['Provider Effective Date'] = merged['VotedDate'].combine_first(merged['Provider Effective Date'])
        merged['Was_Updated'] = False

        orig_dates = df2_original_dates.set_index('Individual NPI')['Provider Effective Date'].to_dict()
        for idx, row in merged.iterrows():
            orig = orig_dates.get(row['Individual NPI'])
            if pd.notna(row['VotedDate']) and pd.isna(orig):
                merged.loc[idx, 'Was_Updated'] = True

        merged.drop(columns=['NPI', 'VotedDate'], inplace=True)

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Merged_Output.xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not save_path:
            window.after(0, lambda: status_label.config(text="Merge cancelled", foreground="orange"))
            reset_gui()
            return

        save_path_global = save_path
        merged.to_excel(save_path, index=False)

        wb = openpyxl.load_workbook(save_path)
        ws = wb.active

        provider_idx = None
        for idx, cell in enumerate(ws[1], 1):
            if cell.value == "Provider Effective Date":
                provider_idx = idx
                break

        if provider_idx is None:
            raise ValueError("Could not find 'Provider Effective Date' column in the output.")

        fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        for df_idx, was_updated in enumerate(merged['Was_Updated']):
            if was_updated:
                excel_row = df_idx + 2
                ws.cell(row=excel_row, column=provider_idx).fill = fill

        for idx, c in enumerate(ws[1], 1):
            if c.value == "Was_Updated":
                ws.delete_cols(idx)
                break

        wb.save(save_path)
        merged_df_global = merged.copy()

        window.after(0, lambda: status_label.config(
            text=f"Merge completed successfully!\nSaved to: {os.path.basename(save_path_global)}", 
            foreground="green"))
        window.after(0, lambda: messagebox.showinfo(
            "Success", 
            f"File successfully saved at:\n{save_path_global}"))
        window.after(0, show_preview)

    except Exception as e:
        error_msg = str(e)
        window.after(0, lambda: status_label.config(text=f"Error: {error_msg}", foreground="red"))
        window.after(0, lambda: messagebox.showerror("Error", error_msg))
    finally:
        window.after(0, reset_gui)

def start_merge_thread():
    threading.Thread(target=perform_merge_logic, daemon=True).start()

def reset_gui():
    merge_button.config(state=tk.DISABLED, background="lightgray")
    select_file1_button.config(state=tk.NORMAL)
    select_file2_button.config(state=tk.NORMAL)
    if not (file1_path_global and file2_path_global):
        status_label.config(text="Please select both files to continue", foreground="orange")

def show_preview():
    global merged_df_global

    for w in preview_frame.winfo_children():
        w.destroy()

    if merged_df_global is None or merged_df_global.empty:
        return

    cols = list(merged_df_global.columns)

    tree = ttk.Treeview(preview_frame, columns=cols, show="headings")
    vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=tree.xview)

    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    vsb.pack(side="right", fill="y")
    hsb.pack(side="bottom", fill="x")
    tree.pack(expand=True, fill="both")

    tree.tag_configure("highlight", background="#FFF9C4")

    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=150, anchor="center")

    for row in merged_df_global.itertuples(index=False):
        values = tuple(row)
        tag = "highlight" if getattr(row, "Was_Updated", False) else ""
        tree.insert("", "end", values=values, tags=(tag,))

# ------------- Professional GUI Design -------------
window = tk.Tk()
window.title("Excel Data Merger - Schedule & Roaster")
window.geometry("1200x800")
window.configure(bg="#f0f0f0")

# Custom style for buttons
style = ttk.Style()
style.theme_use('clam')

# Header
header = tk.Frame(window, bg="#0078D7", height=60)
header.pack(fill="x")

title = tk.Label(header, 
                text="Excel Data Merger - Schedule & Roaster", 
                font=("Segoe UI", 16, "bold"), 
                bg="#0078D7", 
                fg="white")
title.pack(pady=15)

# Main content
content = tk.Frame(window, bg="#f0f0f0")
content.pack(expand=True, fill="both", padx=20, pady=10)

# File selection panel
file_panel = tk.LabelFrame(content, 
                         text=" File Selection ",
                         font=("Segoe UI", 11, "bold"),
                         bg="#f0f0f0",
                         padx=10,
                         pady=10)
file_panel.pack(fill="x", pady=(0, 15))

# File 1 selection
file1_frame = tk.Frame(file_panel, bg="#f0f0f0")
file1_frame.pack(fill="x", pady=5)

select_file1_button = tk.Button(file1_frame, 
                              text="Select Schedule File",
                              font=("Segoe UI", 10),
                              bg="#E1E1E1",
                              relief="groove",
                              command=select_file1)
select_file1_button.pack(side="left", padx=(0, 10))

file1_status = tk.Label(file1_frame, 
                       text="No file selected", 
                       bg="#f0f0f0", 
                       font=("Segoe UI", 10))
file1_status.pack(side="left")

# File 2 selection
file2_frame = tk.Frame(file_panel, bg="#f0f0f0")
file2_frame.pack(fill="x", pady=5)

select_file2_button = tk.Button(file2_frame, 
                              text="Select Roaster File",
                              font=("Segoe UI", 10),
                              bg="#E1E1E1",
                              relief="groove",
                              command=select_file2)
select_file2_button.pack(side="left", padx=(0, 10))

file2_status = tk.Label(file2_frame, 
                       text="No file selected", 
                       bg="#f0f0f0", 
                       font=("Segoe UI", 10))
file2_status.pack(side="left")

# Action buttons
button_frame = tk.Frame(file_panel, bg="#f0f0f0")
button_frame.pack(fill="x", pady=(15, 5))

merge_button = tk.Button(button_frame, 
                       text="Merge Files", 
                       font=("Segoe UI", 10, "bold"),
                       bg="lightgray",
                       fg="white",
                       state=tk.DISABLED,
                       relief="groove",
                       command=start_merge_thread)
merge_button.pack(side="left", padx=(0, 10))

exit_button = tk.Button(button_frame, 
                      text="Exit", 
                      font=("Segoe UI", 10),
                      bg="#E1E1E1",
                      relief="groove",
                      command=window.destroy)
exit_button.pack(side="left")

# Status bar
status_frame = tk.Frame(content, bg="#f0f0f0")
status_frame.pack(fill="x", pady=(0, 15))

status_label = tk.Label(status_frame, 
                      text="Please select both files to continue", 
                      bg="#f0f0f0", 
                      font=("Segoe UI", 10),
                      fg="orange")
status_label.pack()

# Preview panel
preview_panel = tk.LabelFrame(content, 
                            text=" Merged Data Preview ",
                            font=("Segoe UI", 11, "bold"),
                            bg="#f0f0f0",
                            padx=10,
                            pady=10)
preview_panel.pack(expand=True, fill="both")

preview_frame = tk.Frame(preview_panel, bg="#f0f0f0")
preview_frame.pack(expand=True, fill="both", padx=5, pady=5)

window.mainloop()