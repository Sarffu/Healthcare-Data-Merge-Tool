import pandas as pd
from tkinter import Tk, filedialog, ttk, messagebox
import tkinter as tk
import openpyxl
from openpyxl.styles import PatternFill
import threading
import os

# Global variables to store file paths
file1_path_global = None
file2_path_global = None
save_path_global = None # To store the path of the merged file

def select_file1():
    global file1_path_global
    file1_path_global = filedialog.askopenfilename(
        title="Select Excel file with NPI & VotedDate",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file1_path_global:
        file1_label.config(text=f"File 1: {os.path.basename(file1_path_global)}")
        check_and_enable_merge_button()

def select_file2():
    global file2_path_global
    file2_path_global = filedialog.askopenfilename(
        title="Select Excel file with Individual NPI & Provider Effective Date",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file2_path_global:
        file2_label.config(text=f"File 2: {os.path.basename(file2_path_global)}")
        check_and_enable_merge_button()

def check_and_enable_merge_button():
    
    if file1_path_global and file2_path_global:
        merge_button.config(state=tk.NORMAL)
    else:
        merge_button.config(state=tk.DISABLED)

def perform_merge_logic():
    """
    Encapsulates the core merging and highlighting logic.
    This function will be called in a separate thread.
    """
    global save_path_global
    status_label.config(text="Processing merge, please wait...")
    merge_button.config(state=tk.DISABLED)
    select_file1_button.config(state=tk.DISABLED)
    select_file2_button.config(state=tk.DISABLED)

    try:
        # Step 2: Read and clean column names
        df1 = pd.read_excel(file1_path_global)
        df2 = pd.read_excel(file2_path_global)

        df1.columns = df1.columns.str.strip()
        df2.columns = df2.columns.str.strip()

        # Step 3: Backup original Provider Effective Date for comparison
        if 'Individual NPI' not in df2.columns or 'Provider Effective Date' not in df2.columns:
            raise ValueError("File 2 must contain 'Individual NPI' and 'Provider Effective Date' columns.")
        
        df2_original_dates = df2[['Individual NPI', 'Provider Effective Date']].copy()

        # Step 4: Merge
        
        if 'NPI' not in df1.columns or 'VotedDate' not in df1.columns:
            raise ValueError("File 1 must contain 'NPI' and 'VotedDate' columns.")

        merged = pd.merge(df2, df1[['NPI', 'VotedDate']], how='left', left_on='Individual NPI', right_on='NPI')
        merged['Provider Effective Date'] = merged['VotedDate'].combine_first(merged['Provider Effective Date'])

        # Step 5: Track which rows were updated
        merged['Was_Updated'] = False # Initialize to False
        
        # Create a dictionary for faster lookup of original dates
        original_dates_dict = df2_original_dates.set_index('Individual NPI')['Provider Effective Date'].to_dict()

        for index, row in merged.iterrows():
            original_date = original_dates_dict.get(row['Individual NPI'])
            if pd.notna(row['VotedDate']) and pd.isna(original_date):
                merged.loc[index, 'Was_Updated'] = True

        # Drop helper columns
        merged.drop(columns=['NPI', 'VotedDate'], inplace=True)

        # Step 6: Save to Excel
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile="Merged_Output.xlsx", filetypes=[("Excel Files", "*.xlsx")]
        )
        if not save_path:
            window.after(0, lambda: status_label.config(text="Save cancelled."))
            window.after(0, reset_gui_state)
            return

        save_path_global = save_path # Store for confirmation message
        merged.to_excel(save_path, index=False)

        # Step 7: Highlight updated cells using openpyxl
        wb = openpyxl.load_workbook(save_path)
        ws = wb.active

        # Find column index for "Provider Effective Date"
        provider_col_idx = None
        for idx, cell in enumerate(ws[1], 1):
            if cell.value == "Provider Effective Date":
                provider_col_idx = idx
                break

        if provider_col_idx is None:
            raise ValueError("Could not find 'Provider Effective Date' column in the output file for highlighting.")

        # Apply yellow fill for updated rows
        fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        
        for df_row_idx, was_updated in enumerate(merged['Was_Updated']):
            if was_updated:
                # Add 2 to convert DataFrame index to Excel row number (0-indexed DF + 1 for header + 1 for actual data row)
                excel_row_idx = df_row_idx + 2 
                ws.cell(row=excel_row_idx, column=provider_col_idx).fill = fill

        # Remove helper column from Excel
        was_updated_col = None
        for idx, cell in enumerate(ws[1], 1):
            if cell.value == "Was_Updated":
                was_updated_col = idx
                break

        if was_updated_col:
            ws.delete_cols(was_updated_col)

        # Save final Excel file
        wb.save(save_path)
        window.after(0, lambda: status_label.config(text=f"Merge successful! File saved at:\n{save_path_global}"))
        window.after(0, lambda: messagebox.showinfo("Merge Complete", f"Merged file with highlighted updates saved at:\n{save_path_global}"))

    except Exception as e:
        window.after(0, lambda: status_label.config(text=f"Error during merge: {e}"))
        window.after(0, lambda: messagebox.showerror("Merge Error", f"An error occurred during the merge process:\n{e}"))
    finally:
        window.after(0, reset_gui_state)

def start_merge_thread():
    """Starts the merge logic in a separate thread to keep the GUI responsive."""
    threading.Thread(target=perform_merge_logic, daemon=True).start()

def reset_gui_state():
    """Resets the GUI elements to their initial state."""
    merge_button.config(state=tk.DISABLED)
    select_file1_button.config(state=tk.NORMAL)
    select_file2_button.config(state=tk.NORMAL)
    global file1_path_global, file2_path_global, save_path_global
    file1_path_global = None
    file2_path_global = None
    save_path_global = None
    file1_label.config(text="File 1: Not selected")
    file2_label.config(text="File 2: Not selected")
    status_label.config(text="Select both Excel files and click 'Perform Merge'.")


# --- GUI Setup ---
window = tk.Tk()
window.title("Excel Data Merger & Highlighter")
window.geometry("700x400")
window.resizable(False, False)

# --- Styling ---
style = ttk.Style()
style.configure("TButton", font=("Helvetica", 10), padding=10)
style.configure("TLabel", font=("Helvetica", 10))
style.configure("TEntry", font=("Helvetica", 10))

# --- Widgets ---
frame = ttk.Frame(window, padding="20")
frame.pack(expand=True, fill='both')

# File 1 Selection
select_file1_button = ttk.Button(frame, text="Select NPI & VotedDate File (File 1)", command=select_file1)
select_file1_button.pack(pady=5)
file1_label = ttk.Label(frame, text="File 1: Not selected")
file1_label.pack(pady=2)

# File 2 Selection
select_file2_button = ttk.Button(frame, text="Select Individual NPI & Provider Effective Date File (File 2)", command=select_file2)
select_file2_button.pack(pady=5)
file2_label = ttk.Label(frame, text="File 2: Not selected")
file2_label.pack(pady=2)

# Merge Button
merge_button = ttk.Button(frame, text="Perform Merge and Highlight", command=start_merge_thread, state=tk.DISABLED)
merge_button.pack(pady=20)

# Status Label
status_label = ttk.Label(frame, text="Select both Excel files and click 'Perform Merge'.", wraplength=600)
status_label.pack(pady=10)

# Initial state check
check_and_enable_merge_button()

window.protocol("WM_DELETE_WINDOW", window.destroy)
window.mainloop()