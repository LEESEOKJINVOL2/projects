import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import os

wb = Workbook()
ws = wb.active

def initialize_workbook(): # reset workbook
    global wb, ws
    wb.close() # close workbook
    wb = Workbook() # new workbook and active sheet
    ws = wb.active

def transpose(): # matrix transpose 
    file_paths = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
    if not file_paths:
        return
    
    output_folder = filedialog.askdirectory() # select csv file directory
    if not output_folder:
        return

    success_files = [] # for message pop-up
    failed_files = []

    for i, file_path in enumerate(file_paths):
        try:
            initialize_workbook() # reset workbook and active sheet
            dataTotal = None # reset dataTotal

            data = pd.read_csv(file_path, encoding = 'utf-8', sep=",", skiprows=[0])

            cols_to_exclude = ['LAGUNAOUT', 'Laguna_Read_CRC_Counter_NV_X', 'Durant ECID', 'DURANT_LOTID_X_X_X_X_X_X_NV_X', 
                            'LAGUNA_LOTID_X_X_X_X_X_X_NV_X', 'Laguna_Write_CRC_Block_NV_X', 'Laguna_Set_CRC_Counter_NV_X', 
                            'Laguna_CRC_Data_Comparison_NV_X', '_DF1CAL_RF_ERR', 'DF1CAL_RF_OTP', '_GS_', 'Laguna_New_CRC_NV_X'] # trash data
            cols = [c for c in data.columns if not any(exclude_text in c for exclude_text in cols_to_exclude)]
            dataTotal = data[cols]

            transformed = dataTotal.transpose() # matrix transform

            file_name = os.path.basename(file_path) # use same file name
            file_name = f"{os.path.splitext(file_name)[0]}_transposed.csv" # add saved file name "_transformed.xlsx"

            save_path = os.path.join(output_folder, file_name)
            transformed.to_csv(save_path, header=None, encoding = 'utf-8', na_rep = '') # save as xlsx file / header=None(delete 0 1 2 3 ... row)
            success_files.append(file_path)

        except Exception as e:
            failed_files.append(file_path)

    if success_files: # success messagebox
        success_message = "Successfully transposed:\n" + "\n".join(success_files)
        messagebox.showinfo("Matrix Transpose", success_message)

    if failed_files: # failed messagebox
        failed_message = "Transposition error:\n" + "\n".join(failed_files)
        messagebox.showerror("Matrix Transpose", failed_message)

def convert(): # extension convert
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    if not file_paths:
        return
    
    output_folder = filedialog.askdirectory()
    if not output_folder:
        return

    success_files = [] # for message pop-up
    failed_files = []

    for file_path in file_paths:
        try:
            initialize_workbook() # reset workbook and active sheet

            data = pd.read_excel(file_path)
            data = data.replace(True, 1) # True -> 1

            file_name = os.path.basename(file_path)
            file_name = f"{os.path.splitext(file_name)[0]}_converted.csv"

            save_path = os.path.join(output_folder, file_name)
            data.to_csv(save_path, index=False, encoding = 'utf-8', na_rep = '') # save as xlsx file, index=None(delete 0 1 2 3 ... col), header yes
            success_files.append(file_path)
        except Exception as e:
            failed_files.append(file_path)

    if success_files: # success messagebox
        success_message = "Successfully converted:\n" + "\n".join(success_files)
        messagebox.showinfo("Extension Conversion", success_message)

    if failed_files: # failed messagebox
        failed_message = "Conversion error:\n" + "\n".join(failed_files)
        messagebox.showerror("Extension Conversion", failed_message)

def analyze(): # failure analyze
    global wb, ws
    file_paths = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")]) # csv form
    if not file_paths:
        return

    output_folder = filedialog.askdirectory()
    if not output_folder:
        return
    
    success_files = [] # for message pop-up
    failed_files = []

    for file_path in file_paths:
        try:
            initialize_workbook() # reset workbook and active sheet
            data = pd.read_csv(file_path)

            ws.append(list(data.columns))

            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            i = 5 # row initialize
            
            for index, row in data.iterrows():
                if row['SerialNumber'] in ['Site', 'TestResult', 'hbin#', 'sbin#', 'Durant ECID', 'startTime', 'stopTime', 
                    'testTime(ms)', 'DeviceName', 'HandlerID', 'TesterID', 'TestPGM', 'LotID', 'ProcessID', 'OperatorID', 
                    'Loadboard#', 'RetestCode', 'ToteID', 'Temperature', 'LBID_X_X_X_X_X_X_X_NV_X', 'HOSTID_X_X_X_X_X_X_X_NV_X', 
                    'WSPS_ENABLE_X_X_X_X_X_X_NV_X', 'LOAD_PMCAL_TABLE_X_X_X_X_X_NV_X', 'LOAD_KGDCAL_TABLE_X_X_X_X_X_NV_X', 
                    'DURANT_LOTID_X_X_X_X_X_X_NV_X', 'LAGUNA_LOTID_X_X_X_X_X_X_NV_X']: # trash data
                    continue # skip
                elif row['SerialNumber'] in ['hbin', 'sbin', 'Barcode_Err', 'Barcode Err']: # add valid string data
                    ws.append(list(row))
                else:
                    row_floats = [float(value) for value in row[1:]] # string to float
                    
                    for k in range(2, len(row_floats)): # fail data load 
                        if not (pd.isna(row_floats[k]) or
                            (pd.isna(row_floats[1]) and pd.isna(row_floats[0])) or
                            ((not pd.isna(row_floats[1])) and (not pd.isna(row_floats[0])) and (row_floats[1] == row_floats[0] == row_floats[k])) or
                            (pd.isna(row_floats[1]) and (row_floats[k] <= row_floats[0])) or
                            (pd.isna(row_floats[0]) and (row_floats[k] >= row_floats[1])) or
                            ((not pd.isna(row_floats[1])) and (not pd.isna(row_floats[0])) and (row_floats[k] >= row_floats[1]) and (row_floats[k] <= row_floats[0]))):
                            
                            ws.append(list(row)) # fail data row
                            
                            """Normal case
                                1: raw data == NA
                                2: both criteria empty(USL & LSL)
                                3: USL == LSL == data
                                4: only USL is empty & data >= LSL (within spec limit)
                                5: only LSL is empty & data <= USL (within spec limit)
                                6: LSL <= data <= USL"""
                            
                            for j in range(2, len(row_floats)): # add color at fail value
                                if not (pd.isna(row_floats[j]) or
                                    (pd.isna(row_floats[1]) and pd.isna(row_floats[0])) or
                                    ((not pd.isna(row_floats[1])) and (not pd.isna(row_floats[0])) and (row_floats[1] == row_floats[0] == row_floats[j])) or
                                    (pd.isna(row_floats[1]) and (row_floats[j] <= row_floats[0])) or
                                    (pd.isna(row_floats[0]) and (row_floats[j] >= row_floats[1])) or
                                    ((not pd.isna(row_floats[1])) and (not pd.isna(row_floats[0])) and (row_floats[j] >= row_floats[1]) and (row_floats[j] <= row_floats[0]))):
                                    cell = ws.cell(row=i, column=j+2) # color coordinate
                                    cell.fill = yellow_fill # fail data
                            i += 1 # next item
                            break # prevent duplication

            file_name = os.path.basename(file_path)
            file_name = f"{os.path.splitext(file_name)[0]}_analyzed.xlsx"
            save_path = os.path.join(output_folder, file_name)

            wb.save(save_path)
            success_files.append(save_path)

        except Exception as e:
            failed_files.append(file_path)

    if success_files: # success messagebox
        success_message = "Successfully analyzed:\n" + "\n".join(success_files)
        messagebox.showinfo("Failure Analysis", success_message)

    if failed_files: # failed messagebox
        failed_message = "Analysis error:\n" + "\n".join(failed_files)
        messagebox.showerror("Failure Analysis", failed_message)

def main():
    root = tk.Tk()
    root.title("ATE Failure Analyzer_2.0.0")
    root.geometry("350x100")

    transpose_button = tk.Button(root, text="Matrix transpose\n(csv to csv)", command=transpose, width=18, height=2) # transpose button
    transpose_button.grid(row=0, column=0, padx=5, pady=5)

    analyze_button = tk.Button(root, text="Failure analysis\n(csv to xlsx file)", command=analyze, width=18, height=2, background='greenyellow') # Analysis button
    analyze_button.grid(row=0, column=1, padx=5, pady=5)

    quit_button = tk.Button(root, text="Quit", command=root.destroy, width=5, height=2) # quit button
    quit_button.grid(row=0, column=2, padx=5, pady=5)

    convert_button = tk.Button(root, text="Extension conversion\n(xlsx to csv)", command=convert, width=18, height=2) # conversion button
    convert_button.grid(row=1, column=0, padx=5, pady=5)

    initials = tk.Label(root, text="RF LAB1", background='gold')
    initials.grid(row=1, column=2, padx=5, pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()

# 20240405 ATE Failur Analyzer 2.0.0