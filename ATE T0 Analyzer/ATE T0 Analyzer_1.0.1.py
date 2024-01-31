import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import os

def convert(): # raw data load
    file_paths = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
    if not file_paths:
        return

    output_folder = filedialog.askdirectory() # select csv file directory
    if not output_folder:
        return

    for file_path in file_paths:
        try:
            data = pd.read_csv(file_path) # option: header=None

            converted = data.transpose()

            file_name = os.path.basename(file_path) # use same file name
            file_name = f"{os.path.splitext(file_name)[0]}_converted.csv" # add saved file name "_converted.csv" 
            save_path = os.path.join(output_folder, file_name)
            
            converted.to_csv(save_path, index=False) # save as csv file

            print(f"Converted and saved matrix from {file_path} to {save_path}")

        except Exception as e:
            print(f"Error converting {file_path}: {e}")

def load(): # load converted file
    file_paths = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
    if not file_paths:
        return

    output_folder = filedialog.askdirectory()
    if not output_folder:
        return

    for file_path in file_paths:
        try:
            data = pd.read_csv(file_path)

            barcode_err_index = data[data['SerialNumber'] == 'Barcode_Err'].index # search "Barcode_Err" in SerialNumber
            if barcode_err_index.empty:
                print(f"Error analyzing {file_path}: 'Barcode_Err' row not found")
                continue

            wb = Workbook()
            ws = wb.active
            ws.append(list(data.columns)) # add 1 columns (SerialNumber)

            barcode_err_row = data.loc[barcode_err_index[0]] # record Barcode_Err: True or False
            ws.append(list(barcode_err_row))

            data_items = data.loc[barcode_err_index[-1] + 1:, :] # select Test contents

            for _, row in data_items.iterrows(): # SerialNumber: string typ, rest: float type
                serial_number = row['SerialNumber']
                numeric_row = pd.to_numeric(row.iloc[1:], errors='coerce')
                ws.append([serial_number] + list(numeric_row))

            file_name = os.path.basename(file_path) # use same file name
            file_name = f"{os.path.splitext(file_name)[0]}_load.xlsx" # add saved file name "_load"
            save_path = os.path.join(output_folder, file_name)

            wb.save(save_path) # save

            print(f"Analyzed and saved data from {file_path} to {save_path}")

        except Exception as e:
            print(f"Error analyzing {file_path}: {e}")

def filter_data(data, lower_limit_col, upper_limit_col):
    filtered_data = pd.DataFrame(columns = data.columns)
    
    for index, row in data.iterrows():
        lower_limit = row['Lower Limit']
        upper_limit = row['Upper Limit']

        if row['SerialNumber'] == 'DURANT_LOTID_X_X_X_X_X_X_NV_X': # string type data
            continue # skip

        non_nan_data = row.iloc[3:].dropna() # extract non_NA data

        if non_nan_data.empty:
            continue # skip

        if pd.isna(lower_limit) and pd.isna(upper_limit): # both empty(USL & LSL)
            continue # skip
        elif (lower_limit == upper_limit) and (non_nan_data.eq(lower_limit).all()): # USL == LSL == data
            continue # skip
        elif pd.isna(upper_limit): # only USL is empty
            if (non_nan_data >= lower_limit).all(): # data >= LSL (within spec limit)
                continue # skip
        elif pd.isna(lower_limit): # only LSL is empty
            if (non_nan_data <= upper_limit).all(): # data <= USL (within spec limit)
                continue # skip
        elif (non_nan_data >= lower_limit).all() and (non_nan_data <= upper_limit).all(): # LSL <= data <= USL
            continue # skip

        filtered_data = filtered_data.append(row) # save failed data columns
    return filtered_data

def analyze():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    if not file_paths:
        return

    output_folder = filedialog.askdirectory()
    if not output_folder:
        return

    for file_path in file_paths:
        try:
            data = pd.read_excel(file_path)

            filtered_data = filter_data(data, "Lower Limit", "Upper Limit")
            
            file_name = os.path.basename(file_path) # use same file name
            file_name = f"{os.path.splitext(file_name)[0]}_filtered_data.xlsx" # add saved file name "_filtered_data"
            save_path = os.path.join(output_folder, file_name)

            filtered_data.to_excel(save_path, index=False)

            print(f"Analyzed and saved filtered data from {file_path} to {save_path}")

        except Exception as e:
            print(f"Error analyzing {file_path}: {e}")

def main():
    root = tk.Tk()
    root.title("ATE T0 Analyzer_1.0.1")
    root.geometry("300x120")

    convert_button = tk.Button(root, text="Convert", command=convert, width=10, height=2) # convert button
    convert_button.pack(side=tk.LEFT, padx=10, pady=10)
    
    quit_button = tk.Button(root, text="Quit", command=root.destroy, width=5, height=2) # quit button
    quit_button.pack(side=tk.RIGHT, padx=10, pady=10)
    
    load_button = tk.Button(root, text="Load", command=load, width=10, height=2) # load button
    load_button.pack(side=tk.TOP, padx=10, pady=10)

    analyze_button = tk.Button(root, text="Analyze", command=analyze, width=10, height=2) # analyze button
    analyze_button.pack(side=tk.TOP, padx=10, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()

# 240130
# To do list
    # 1 convert button
    #   remove 1st column(numbering 0 1 2 3 4) at 1st file
    # 2 filtering
    #   add color when data was failed  