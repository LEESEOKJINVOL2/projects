import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import pandas as pd
import sys 
import datetime
from typing import Union
import matplotlib.pyplot as plt
import seaborn as sns 
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR 

PROCESSING_ITEMS = {
    'CWGSHP': 'CWGSHP', 
    'CWGSLP': 'CWGSLP'
}

# Adjusted size for a 2x2 layout based on a landscape A4 slide
PPTX_IMG_W = Inches(4.5)  # Image Width (Horizontal)
PPTX_IMG_H = Inches(3.0)  # Image Height (Vertical)
PPTX_GAP = Inches(0.2)    # Gap between images

# 3. PPTX Layout Configuration
PPTX_TITLE_MARGIN_TOP = Inches(1.5) # Top margin occupied by the slide title area
PPTX_FILENAME_LABEL_H = Inches(0.2) # Height occupied by the filename label

SLIDE_PLAN = [
    # 2. CWPRLP_RF_PWR
    {
        'title': "CWPRLP_RF_PWR",
        'files': [
            "CWPRLP_RF_PWR_730.png",
            "CWPRLP_RF_PWR_1200.png",
            "CWPRLP_RF_PWR_1500.png"
        ]
    },
    # 3. CWPRHP_RF_PWR
    {
        'title': "CWPRHP_RF_PWR",
        'files': [
            "CWPRHP_RF_PWR_730.png",
            "CWPRHP_RF_PWR_1500.png"
        ]
    },
    # 4. CWGSLP_RF_PWR_730
    {
        'title': "CWGSLP_RF_PWR_730",
        'files': [
            "TXFE0_CWGSLP_RF_PWR_CH039_730.png",
            "TXFE2_CWGSLP_RF_PWR_CH063_730.png",
            "TXFE3_CWGSLP_RF_PWR_CH038_730.png",
            "TXFE4_CWGSLP_RF_PWR_CH063_730.png"
        ]
    },
    # 5. CWGSLP_RF_PWR_730
    {
        'title': "CWGSLP_RF_PWR_730",
        'files': [
            "TXFE5_CWGSLP_RF_PWR_CH062_730.png",
            "TXFE6_CWGSLP_RF_PWR_CH062_730.png",
            "TXFE7_CWGSLP_RF_PWR_CH062_730.png"
        ]
    },
    # 6. CWGSLP_RF_PWR_1200
    {
        'title': "CWGSLP_RF_PWR_1200",
        'files': [
            "TXFE0_CWGSLP_RF_PWR_CH039_1200.png",
            "TXFE2_CWGSLP_RF_PWR_CH063_1200.png",
            "TXFE3_CWGSLP_RF_PWR_CH038_1200.png",
            "TXFE4_CWGSLP_RF_PWR_CH063_1200.png"
        ]
    },
    # 7. CWGSLP_RF_PWR_1200
    {
        'title': "CWGSLP_RF_PWR_1200",
        'files': [
            "TXFE5_CWGSLP_RF_PWR_CH062_1200.png",
            "TXFE6_CWGSLP_RF_PWR_CH062_1200.png",
            "TXFE7_CWGSLP_RF_PWR_CH062_1200.png"
        ]
    },
    # 8. MINAGAIN_RF_GAIN_M25_MIX
    {
        'title': "MINAGAIN_RF_GAIN_M25_MIX",
        'files': [
            "MINAGAIN_RF_GAIN_M25_BW1P5_MIX.png",
            "MINAGAIN_RF_GAIN_M25_BW2P0_MIX.png",
            "MINAGAIN_RF_GAIN_M25_BW3P0_MIX.png"
        ]
    },
    # 9. MAXAGAIN_RF_GAIN_M75_MIX
    {
        'title': "MAXAGAIN_RF_GAIN_M75_MIX",
        'files': [
            "MAXAGAIN_RF_GAIN_M75_BW1P5_MIX.png",
            "MAXAGAIN_RF_GAIN_M75_BW2P0_MIX.png",
            "MAXAGAIN_RF_GAIN_M75_BW3P0_MIX.png"
        ]
    },
    # 10. XAGAIN_RF_GAIN_M75_LNA
    {
        'title': "MAXAGAIN_RF_GAIN_M75_LNA",
        'files': [
            "MAXAGAIN_RF_GAIN_M75_BW1P5_LNA.png",
            "MAXAGAIN_RF_GAIN_M75_BW2P0_LNA.png",
            "MAXAGAIN_RF_GAIN_M75_BW3P0_LNA.png"
        ]
    },
    # 11. PALDO_RF_PWR_LP_0_730
    {
        'title': "PALDO_RF_PWR_LP_0_730",
        'files': [
            "PALDO_RF_PWR_LP_0_730_0.png",
            "PALDO_RF_PWR_LP_0_730_512.png",
            "PALDO_RF_PWR_LP_0_730_2048.png"
        ]
    },
    # 12. PALDO_RF_PWR_LP_0_1200
    {
        'title': "PALDO_RF_PWR_LP_0_1200",
        'files': [
            "PALDO_RF_PWR_LP_0_1200_0.png",
            "PALDO_RF_PWR_LP_0_1200_1024.png",
            "PALDO_RF_PWR_LP_0_1200_4095.png"
        ]
    },
    # 13. PALDO_RF_PWR_LP_0_1500
    {
        'title': "PALDO_RF_PWR_LP_0_1500",
        'files': [
            "PALDO_RF_PWR_LP_0_1500_0.png",
            "PALDO_RF_PWR_LP_0_1500_1024.png",
            "PALDO_RF_PWR_LP_0_1500_4095.png"
        ]
    },
    # 14. PALDO_RF_PWR_LP_13_730
    {
        'title': "PALDO_RF_PWR_LP_13_730",
        'files': [
            "PALDO_RF_PWR_LP_13_730_0.png",
            "PALDO_RF_PWR_LP_13_730_512.png",
            "PALDO_RF_PWR_LP_13_730_2048.png"
        ]
    },
    # 15. PALDO_RF_PWR_LP_13_1200
    {
        'title': "PALDO_RF_PWR_LP_13_1200",
        'files': [
            "PALDO_RF_PWR_LP_15_1200_0.png",
            "PALDO_RF_PWR_LP_15_1200_1024.png",
            "PALDO_RF_PWR_LP_15_1200_4095.png"
        ]
    },
    # 16. PALDO_RF_PWR_LP_13_1500
    {
        'title': "PALDO_RF_PWR_LP_13_1500",
        'files': [
            "PALDO_RF_PWR_LP_15_1500_0.png",
            "PALDO_RF_PWR_LP_15_1500_1024.png",
            "PALDO_RF_PWR_LP_15_1500_4095.png"
        ]
    },
    # 17. CW_RF_RSSIERR_RXFE0_CH039
    {
        'title': "CW_RF_RSSIERR_RXFE0_CH039",
        'files': [
            "CW_RF_RSSIERR_RXFE0_CH039_BW1P5.png",
            "CW_RF_RSSIERR_RXFE0_CH039_BW2P0.png",
            "CW_RF_RSSIERR_RXFE0_CH039_BW3P0.png"
        ]
    },
    # 18. CW_RF_RSSIERR_RXFE1_CH050
    {
        'title': "CW_RF_RSSIERR_RXFE1_CH050",
        'files': [
            "CW_RF_RSSIERR_RXFE1_CH050_BW1P5.png",
            "CW_RF_RSSIERR_RXFE1_CH050_BW2P0.png",
            "CW_RF_RSSIERR_RXFE1_CH050_BW3P0.png"
        ]
    },
    # 19. CW_RF_RSSIERR_RXFE2_CH063
    {
        'title': "CW_RF_RSSIERR_RXFE2_CH063",
        'files': [
            "CW_RF_RSSIERR_RXFE2_CH063_BW1P5.png",
            "CW_RF_RSSIERR_RXFE2_CH063_BW2P0.png",
            "CW_RF_RSSIERR_RXFE2_CH063_BW3P0.png"
        ]
    },
    # 20. CW_RF_RSSIERR_RXFE4_CH063
    {
        'title': "CW_RF_RSSIERR_RXFE4_CH063",
        'files': [
            "CW_RF_RSSIERR_RXFE4_CH063_BW1P5.png",
            "CW_RF_RSSIERR_RXFE4_CH063_BW2P0.png",
            "CW_RF_RSSIERR_RXFE4_CH063_BW3P0.png"
        ]
    },
    # 21. CW_RF_RSSIERR_RXFE5_CH062
    {
        'title': "CW_RF_RSSIERR_RXFE5_CH062",
        'files': [
            "CW_RF_RSSIERR_RXFE5_CH062_BW1P5.png",
            "CW_RF_RSSIERR_RXFE5_CH062_BW2P0.png",
            "CW_RF_RSSIERR_RXFE5_CH062_BW3P0.png"
        ]
    },
    # 22. CW_RF_RSSIERR_RXFE6_CH062
    {
        'title': "CW_RF_RSSIERR_RXFE6_CH062",
        'files': [
            "CW_RF_RSSIERR_RXFE6_CH062_BW1P5.png",
            "CW_RF_RSSIERR_RXFE6_CH062_BW2P0.png",
            "CW_RF_RSSIERR_RXFE6_CH062_BW3P0.png"
        ]
    },
    # 23. CW_RF_RSSIERR_RXFE7_CH062
    {
        'title': "CW_RF_RSSIERR_RXFE7_CH062",
        'files': [
            "CW_RF_RSSIERR_RXFE7_CH062_BW1P5.png",
            "CW_RF_RSSIERR_RXFE7_CH062_BW2P0.png",
            "CW_RF_RSSIERR_RXFE7_CH062_BW3P0.png"
        ]
    }
]

FIXED_COLUMNS = [
    'TEST_NAME', 
    'LLIM', 
    'ULIM', 
    'BLOCK', 
    'CH', 
    'PWR', 
    'FREQ/VOLT',
    'OFFS' 
]

RESULT_COLUMN_NAME = 'RESULT'

FINAL_EXCEL_FILENAME = 'Combined_Test_Results.xlsx'

# The Keys in this dictionary will become the sheet names in the Excel file.
PROCESSING_ITEMS = {
    'CWGSLP': 'CWGSLP_RF_PWR',
    'CWGSHP': 'CWGSHP_RF_PWR',
    'CWPRLP': 'CWPRLP_RF_PWR',
    'CWPRHP': 'CWPRHP_RF_PWR',
    'MAXA GAIN': 'MAXAGAIN_RF_GAIN',
    'MINA GAIN': 'MINAGAIN_RF_GAIN',
    'PALDO PWR': 'PALDO_RF_PWR',
    'CW RSSIERR': 'CW_RF_RSSIERR'
}

# Configuration for plotting
PLOT_X_AXIS_LABELS = ['BLOCK', 'CH', 'PWR'] 
LIMIT_COLUMNS = ['LLIM', 'ULIM']
# The grouping criteria for plots have been changed to two: BLOCK and FREQ/VOLT
GROUPING_COLUMNS_FOR_PLOTS = ['FREQ/VOLT', 'BLOCK'] 

# Global variables
all_data = {}
root_path = None 
log_message = None
results = None
root = None

def select_folder_and_process_csv():
    global all_data, root_path
    
    # 1. Select the specific folder (root folder)
    new_root_folder_path = filedialog.askdirectory(
        title="Select the root folder containing subfolders with CSV files."
    )

    # Check if the user cancelled the folder selection
    if not new_root_folder_path:
        log_message.set("Folder selection cancelled.")
        results.yview_moveto(1.0)
        return

    root_path = new_root_folder_path # Store the selected root path
    
    # Initialize results and global data
    results.delete('1.0', tk.END)
    log_message.set(f"Selected Root Folder: {root_path}")
    all_data = {}
    csv_file_count = 0
    subfolder_count = 0

    try:
        entries = os.listdir(root_path)

        for entry in entries:
            subfolder_path = os.path.join(root_path, entry)

            if os.path.isdir(subfolder_path):
                subfolder_count += 1
                subfolder_name = entry
                subfolder_dfs = []
                found_in_subfolder = False

                results.insert(tk.END, f"\n--- Subfolder: {subfolder_name} ---\n", 'subfolder')
                subfolder_files = os.listdir(subfolder_path)

                for file_name in subfolder_files:
                    if file_name.endswith('.csv'):
                        csv_file_path = os.path.join(subfolder_path, file_name)

                        try:
                            df = pd.read_csv(csv_file_path)
                            subfolder_dfs.append(df)
                            results.insert(
                                tk.END, 
                                f"  > Loaded: {file_name} (Shape: {df.shape})\n"
                            )
                            csv_file_count += 1
                            found_in_subfolder = True
                        
                        except Exception as e:
                            results.insert(
                                tk.END, 
                                f"  > Error reading {file_name}: {e}\n", 
                                'error'
                            )
                
                if subfolder_dfs:
                    combined_df = pd.concat(subfolder_dfs, ignore_index=True)
                    all_data[subfolder_name] = combined_df
                
                if not found_in_subfolder:
                    results.insert(tk.END, "  > No CSV files found in this subfolder.\n", 'not_found')

        if csv_file_count > 0 and subfolder_count > 0:
            run_all_processing_and_save()
            results.yview_moveto(1.0)
            log_message.set(f"Load Complete! {csv_file_count} files from {subfolder_count} folders loaded. Ready to run ALL processing.")
            
        else:
            results.yview_moveto(1.0)
            log_message.set("No valid data found in the selected folder structure.")
            
    except Exception as e:
        log_message.set(f"An unexpected error occurred: {e}")
        results.yview_moveto(1.0)

def _process_single_item(sheet_name: str, filter_string: str) -> Union[pd.DataFrame, None]:
    if not all_data:
        return None
    
    processed_dfs = []
    base_df = None
    
    results.insert(tk.END, f"\n\n--- Processing Item: {sheet_name} (Filter: '{filter_string}') ---\n", 'subfolder')

    for subfolder_name, df in all_data.items():
        try:
            # 1. Filter by TEST_NAME containing the filter string
            filtered_df = df[df['TEST_NAME'].astype(str).str.contains(filter_string, na=False, case=False)].copy()
        except KeyError:
             results.insert(
                 tk.END, 
                 f"  > ERROR: Subfolder {subfolder_name} is missing 'TEST_NAME' column. Skipping.\n", 
                 'error'
               )
             continue

        if filtered_df.empty:
            results.insert(tk.END, f"  > Subfolder {subfolder_name}: No matching data found. Skipping.\n")
            continue
        
        # Check if the required columns exist
        required_cols = FIXED_COLUMNS + [RESULT_COLUMN_NAME]
        missing_cols = [col for col in required_cols if col not in filtered_df.columns]
        if missing_cols:
             results.insert(
                 tk.END, 
                 f"  > ERROR: Subfolder {subfolder_name} is missing required columns: {', '.join(missing_cols)}. Skipping.\n", 
                 'error'
               )
             continue
        
        # Select the fixed columns and reset index
        fixed_part = filtered_df[FIXED_COLUMNS].reset_index(drop=True)
        # Select the result column
        result_part = filtered_df[RESULT_COLUMN_NAME].reset_index(drop=True)
        
        # 3. Rename the RESULT column to the folder name
        result_part.name = subfolder_name 
        
        processed_dfs.append(result_part)
        
        # 4. Use the first processed fixed part as the base DataFrame
        if base_df is None:
            base_df = fixed_part
        
    # Final merge logic
    if base_df is None or not processed_dfs:
        results.insert(tk.END, f"  > WARNING: Could not generate data for sheet '{sheet_name}'.\n", 'error')
        return None

    # Concatenate the base fixed columns with all the processed result columns side-by-side
    final_df = pd.concat([base_df] + processed_dfs, axis=1)
    
    results.insert(tk.END, f"  > SUCCESS: Final DataFrame generated. Shape: {final_df.shape}\n", 'saved')
    return final_df

def run_all_processing_and_save():
  global root_path

  if not all_data:
    messagebox.showwarning("Warning", "No data loaded yet. Please select and process a folder first.")
    return

  # Handle case where root_path is not set (e.g., select_folder_and_process_csv was skipped)
  if root_path is None:
    messagebox.showerror("Error", "Root directory path is not set. Please load data first.")
    log_message.set("Error: Root directory path is not set.")
    results.yview_moveto(1.0)
    return

  # 1. Generate timestamped filename (YYYYMMDD_HHMMSS)
  now = datetime.datetime.now()
  timestamp = now.strftime("%Y%m%d_%H%M%S")
  
  # Filename format: result sheet_YYYYMMDD_HHMMSS.xlsx
  dynamic_filename = f"result sheet_{timestamp}.xlsx"
  
  # 2. Construct the full file path in the root directory
  excel_file_path = os.path.join(root_path, dynamic_filename)

  log_message.set(f"Running batch processing and saving to {excel_file_path}...")
  results.insert(tk.END, "\n\n--- STARTING BATCH SAVE TO EXCEL ---\n", 'subfolder')
  
  successful_sheets = []

  try:
    # Create a pandas ExcelWriter object
    with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
      
      for sheet_name, filter_string in PROCESSING_ITEMS.items():
        
        final_df = _process_single_item(sheet_name, filter_string)
        
        if final_df is not None:
          final_df.to_excel(writer, sheet_name=sheet_name, index=False)
          results.insert(tk.END, f" > Sheet '{sheet_name}' saved successfully.\n", 'saved')
          successful_sheets.append(sheet_name)
        else:
          results.insert(tk.END, f" > Sheet '{sheet_name}' skipped due to no data.\n", 'error')

    if successful_sheets:
      log_message.set(f"Batch Save Complete! {len(successful_sheets)} sheets saved to {excel_file_path}.")
      results.yview_moveto(1.0)
      messagebox.showinfo("Success", f"All data sheets have been saved to:\n{excel_file_path}\nSheets saved: {', '.join(successful_sheets)}")
    else:
      log_message.set(f"Batch Save Finished, but NO sheets were saved.")
      results.yview_moveto(1.0)
      messagebox.showwarning("Warning", "No data matched any defined filter criteria.")

  except Exception as e:
    log_message.set(f"Error during Excel saving: {e}")
    results.yview_moveto(1.0)
    messagebox.showerror("Error", f"Failed to save Excel file:\n{e}")

def generate_plots():
    global root_path
    
    # 1. Select the Excel file path
    initial_dir = root_path if root_path else os.getcwd()
    excel_file_path = filedialog.askopenfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialdir=initial_dir,
        title="Select the generated combined Excel file to plot from"
    )
    
    if not excel_file_path:
        log_message.set("Excel file selection cancelled for plotting.")
        results.yview_moveto(1.0)
        return
    
    plot_base_dir = os.path.dirname(excel_file_path)
    
    log_message.set(f"Reading ALL sheets from '{os.path.basename(excel_file_path)}' for plotting...")
    results.yview_moveto(1.0)
    results.insert(tk.END, f"\n\n--- STARTING BATCH PLOT GENERATION ---\n", 'subfolder')

    try:
        # Read all sheets into a Dictionary
        all_sheets_data = pd.read_excel(excel_file_path, sheet_name=None)
        
    except ValueError:
        messagebox.showerror(
            "Error: Invalid File Format", 
            f"The selected file format is invalid.\n\n[3. Generate Plots] requires a single **.xlsx** file (Excel Workbook)."
        )
        log_message.set(f"Plot generation failed: Invalid file format.")
        results.yview_moveto(1.0)
        return
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred while reading the Excel file: {e}")
        log_message.set(f"Plot generation failed: {e}")
        results.yview_moveto(1.0)
        return

    # Create a folder for plots 
    plots_folder = os.path.join(plot_base_dir, "Plots")
    os.makedirs(plots_folder, exist_ok=True)
    results.insert(tk.END, f"  > Saving plots to: {plots_folder}\n")

    total_plot_count = 0
    successful_sheet_count = 0
    
    # Define sheet groups for different logic
    PR_SHEETS = ['CWPRLP', 'CWPRHP']
    GAIN_SHEETS = ['MAXA GAIN', 'MINA GAIN'] 
    PALDO_SHEETS = ['PALDO PWR'] 
    RSSIERR_SHEETS = ['CW RSSIERR'] 
    
    CW_POWER_SHEETS = ['CWGSLP', 'CWGSHP'] 
    
    # Iterate through all sheets in the Dictionary to generate plots.
    for sheet_name, df in all_sheets_data.items():
        results.insert(tk.END, f"\n--- Processing Sheet: {sheet_name} ---\n", 'subfolder')
        
        # Default (includes CWGSLP, CWGSHP initially)
        group_cols = GROUPING_COLUMNS_FOR_PLOTS # ['FREQ/VOLT', 'BLOCK']
        sort_cols = PLOT_X_AXIS_LABELS # ['BLOCK', 'CH', 'PWR']
        group_filter = None 
        ascending_flags = [True] * len(sort_cols) 
        plot_limits = True # Whether to plot the limits

        if sheet_name in CW_POWER_SHEETS:
            results.insert(tk.END, f"  > NOTE: Sheet '{sheet_name}' is a CW Power sheet. Plotting limits (LLIM/ULIM) disabled.\n")
            plot_limits = False
            
        elif sheet_name in PR_SHEETS:
            # CWPRLP / CWPRHP Logic:
            group_cols = ['FREQ/VOLT'] 
            sort_cols = ['BLOCK', 'CH']
            ascending_flags = [True, True]
            
        elif sheet_name in GAIN_SHEETS:
            # MAXAGAIN / MINAGAIN Logic:
            group_cols = ['OFFS', 'FREQ/VOLT'] 
            sort_cols = ['OFFS', 'BLOCK', 'CH']
            group_filter = r'(LNA|MIX)' 
            ascending_flags = [True, True, True] 
        
        elif sheet_name in PALDO_SHEETS:
            # PALDO PWR Logic:
            group_cols = ['CH', 'PWR', 'FREQ/VOLT', 'OFFS'] 
            sort_cols = ['BLOCK', 'FREQ/VOLT'] 
            group_filter = r'(PWR|OFFS)' 
            ascending_flags = [True, True] 
            plot_limits = False 
            
        # [FIX] CW RSSIERR Logic added
        elif sheet_name in RSSIERR_SHEETS:
            # CW RSSIERR Logic:
            group_cols = ['BLOCK', 'CH', 'OFFS'] # Grouping: BLOCK, CH, OFFS (Request)
            sort_cols = ['PWR'] # X-axis: PWR (Request)
            ascending_flags = [True] # PWR ascending (RSSIERR)
            group_filter = None 
            plot_limits = True 

        # If it's not the RSSIERR sheet and 'PWR' is in the sort columns, set it to descending.
        if sheet_name not in RSSIERR_SHEETS and 'PWR' in sort_cols: 
            pwr_index = sort_cols.index('PWR')
            ascending_flags[pwr_index] = False # Descending (False)

        required_cols_check = sort_cols + group_cols
        if plot_limits:
            required_cols_check += LIMIT_COLUMNS
        
        if not all(col in df.columns for col in required_cols_check):
            results.insert(tk.END, f"  > WARNING: Sheet '{sheet_name}' skipped. Missing required columns (e.g., {', '.join(required_cols_check)}).\n", 'error')
            continue

        try:
            # 1. Drop 'TEST_NAME' if they exist
            cols_to_drop = [col for col in ['TEST_NAME'] if col in df.columns]
            if cols_to_drop:
                df = df.drop(columns=cols_to_drop)
            # Convert BLOCK, CH, OFFS, PWR, etc. to strings to prevent errors due to NaN or strange strings
            cols_to_str = ['BLOCK', 'CH', 'FREQ/VOLT', 'OFFS', 'PWR'] 
            for col in cols_to_str:
                if col in df.columns:
                    # Replace NaN with empty string, then convert to string and strip whitespace
                    df[col] = df[col].astype(str).str.strip().fillna('') 

            # 3. [FIX] Create a temporary column for numeric PWR sorting and replace the sort column
            sort_cols_for_execution = list(sort_cols)
            
            # Check if PWR is a sort target
            is_pwr_sort_target = 'PWR' in sort_cols
            
            if is_pwr_sort_target:
                pwr_index = sort_cols.index('PWR')
                is_pwr_desc_sort = ascending_flags[pwr_index] == False # Check if descending sort
                
                # Use errors='coerce' to turn values that can't be converted to numbers into NaN.
                df['PWR_NUMERIC_SORT'] = pd.to_numeric(df['PWR'], errors='coerce') 
                
                # Handle NaN values: set them to be pushed to the end based on the sorting intention
                if is_pwr_desc_sort: # Descending (Largest value first, NaN last)
                    # Set NaN to -infinity to push them to the end
                    df['PWR_NUMERIC_SORT'] = df['PWR_NUMERIC_SORT'].fillna(-float('inf'))
                else: # Ascending (Smallest value first, NaN last)
                    # Set NaN to +infinity to push them to the end
                    df['PWR_NUMERIC_SORT'] = df['PWR_NUMERIC_SORT'].fillna(float('inf'))

                # Replace the 'PWR' column in the execution list with the temporary numeric column
                pwr_sort_index = sort_cols_for_execution.index('PWR')
                sort_cols_for_execution[pwr_sort_index] = 'PWR_NUMERIC_SORT'
                results.insert(tk.END, f"  > NOTE: Sheet '{sheet_name}' using 'PWR_NUMERIC_SORT' for stable sorting.\n")

            # This is left here for filtering on columns other than TEST_NAME (e.g., 'FREQ/VOLT' filtering for GAIN sheets)
            if group_filter and sheet_name in GAIN_SHEETS:
                df = df[df['FREQ/VOLT'].astype(str).str.contains(group_filter, na=False, case=False)].copy()
                if df.empty:
                    results.insert(tk.END, f"  > WARNING: Sheet '{sheet_name}' skipped after filtering by '{group_filter}'. No data remaining.\n", 'error')
                    continue

            # [FIX] Add PWR_NUMERIC_SORT to fixed_cols_temp to exclude it from data columns
            fixed_cols_temp = [col for col in FIXED_COLUMNS if col in df.columns] 
            if plot_limits:
                fixed_cols_temp += LIMIT_COLUMNS
            if 'PWR_NUMERIC_SORT' in df.columns:
                 fixed_cols_temp.append('PWR_NUMERIC_SORT')
                 
            data_columns = [col for col in df.columns if col not in fixed_cols_temp]
            
            if not data_columns:
                results.insert(tk.END, "  > WARNING: No data columns found to plot. Skipping sheet.\n", 'error')
                continue

            grouped = df.groupby(group_cols)
            plot_count_in_sheet = 0
            
            for group_key, group_df in grouped:
                
                # Skip groups with 1 or less rows (no meaning for a plot)
                if len(group_df) < 2:
                    continue

                # [FIX] Modified all group key extraction logic, including RSSIERR logic
                group_block, group_ch, group_pwr, group_freq_volt, group_offs = '', 'NOCH', 'NOPWR', 'NOFREQ', 'NOOFFS'

                if sheet_name in PALDO_SHEETS:
                    group_ch, group_pwr, group_freq_volt, group_offs = group_key
                elif sheet_name in RSSIERR_SHEETS: # ⬅️ RSSIERR Group Key Extraction
                    group_block, group_ch, group_offs = group_key
                    # Use the first value for PWR and FREQ/VOLT within the group
                    group_pwr = str(group_df['PWR'].iloc[0]) 
                    # group_freq_volt = str(group_df['FREQ/VOLT'].iloc[0]) 
                elif sheet_name in PR_SHEETS:
                    group_freq_volt = group_key
                    group_pwr = str(group_df['PWR'].iloc[0]) 
                    group_ch = str(group_df['CH'].iloc[0]) 
                elif sheet_name in GAIN_SHEETS:
                    group_offs, group_freq_volt = group_key 
                    group_pwr = str(group_df['PWR'].iloc[0]) 
                    group_ch = str(group_df['CH'].iloc[0]) 
                else: # Default (CWGSLP, CWGSHP, and others)
                    group_freq_volt, group_block = group_key 
                    group_pwr = str(group_df['PWR'].iloc[0]) 
                    group_ch = str(group_df['CH'].iloc[0]) 
                    
                # [FIX] Use sort_cols_for_execution (which may include PWR_NUMERIC_SORT)
                group_df = group_df.sort_values(by=sort_cols_for_execution, ascending=ascending_flags).reset_index(drop=True)
                
                # Create X-axis labels (using defined sort_cols - original PWR string for display)
                group_df['X_LABEL'] = group_df[sort_cols].astype(str).agg(' '.join, axis=1)
                
                clean_freq_volt = str(group_freq_volt).replace('/', '_')
                clean_offs = str(group_offs).replace('/', '_')
                clean_pwr = str(group_pwr).replace('/', '_')
                test_name_part = PROCESSING_ITEMS[sheet_name]
                
                
                if sheet_name in PALDO_SHEETS:
                    # PALDO Filename: PALDO_RF_PWR_CH_PWR_FREQ/VOLT_OFFS.png
                    png_title = f"{test_name_part}_{group_ch}_{clean_pwr}_{clean_freq_volt}_{clean_offs}"
                    
                elif sheet_name in RSSIERR_SHEETS: # ⬅️ RSSIERR Title Generation
                    # RSSIERR Filename: CW_RF_RSSIERR_BLOCK_CH_OFFS.png
                    png_title = f"{test_name_part}_{group_block}_{group_ch}_{clean_offs}" 
                    
                elif sheet_name in GAIN_SHEETS:
                    # MAXAGAIN/MINAGAIN Filename: TESTNAME_PWR_OFFS_FREQ/VOLT.png
                    png_title = f"{test_name_part}_{clean_pwr}_{clean_offs}_{clean_freq_volt}" 
                    
                elif sheet_name in PR_SHEETS:
                    # CWPRLP / CWPRHP Filename: TESTNAME_FREQ_VOLT.png 
                    png_title = f"{test_name_part}_{clean_freq_volt}"
                    
                else:
                    # Default Filename (Includes CWGSLP/CWGSHP): BLOCK_TESTNAME_CH_FREQ_VOLT.png
                    png_title = f"{group_block}_{test_name_part}_{group_ch}_{clean_freq_volt}"
                    
                plot_filename = os.path.join(
                    plots_folder, 
                    f"{png_title}.png"
                )

                # 6. Plotting 
                plt.figure(figsize=(15, 8)) 
                sns.set_style("whitegrid") 
                
                # Plot LLIM and ULIM (Only if plot_limits is True)
                if plot_limits:
                    plt.plot(group_df['X_LABEL'], group_df[LIMIT_COLUMNS[0]], marker='', linestyle='-', label=f"{LIMIT_COLUMNS[0]}", color='steelblue')
                    plt.plot(group_df['X_LABEL'], group_df[LIMIT_COLUMNS[1]], marker='', linestyle='-', label=f"{LIMIT_COLUMNS[1]}", color='firebrick')

                # Plot each data column (No marker)
                for col in data_columns:
                    if "Ref" in col:
                        plt.plot(group_df['X_LABEL'], group_df[col], marker='', linestyle='--', color='red', label=f"{col}")
                    else:
                        plt.plot(group_df['X_LABEL'], group_df[col], marker='', linestyle='-', label=f"{col}")
                
                plt.title(png_title, fontsize=24)
                
                # Adjust x-tick rotation based on the number of labels
                if len(group_df['X_LABEL']) > 50:
                    rotation = 90
                    fontsize = 6
                else:
                    rotation = 60
                    fontsize = 8
                    
                plt.xticks(rotation=rotation, fontsize=fontsize) 
                plt.yticks(fontsize=16)
                
                # Legend outside top right
                plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', borderaxespad=0)
                plt.tight_layout() 

                # Save the plot
                plt.savefig(plot_filename)
                plt.close() 

                results.insert(tk.END, f"  > Plot saved: {os.path.basename(plot_filename)}\n")
                plot_count_in_sheet += 1
                total_plot_count += 1
            
            if plot_count_in_sheet > 0:
                successful_sheet_count += 1

        except Exception as e:
            results.insert(tk.END, f"  > ERROR processing sheet '{sheet_name}': {e}. Skipping.\n", 'error')

    if total_plot_count > 0:
        log_message.set(f"Plot generation complete! {total_plot_count} plots from {successful_sheet_count} sheets saved to '{plots_folder}'.")
        results.yview_moveto(1.0)
        messagebox.showinfo("Success", f"Plot generation complete! {total_plot_count} plots saved.")
    else:
        log_message.set("Plot generation finished. No plots were generated.")
        results.yview_moveto(1.0)
        messagebox.showwarning("Warning", "No plots were generated. Check if the Excel data contains required columns.")

def generate_ppt():
    global root_path
    global log_message, results
    global SLIDE_PLAN, PPTX_IMG_W, PPTX_IMG_H, PPTX_GAP
    global PPTX_TITLE_MARGIN_TOP
    
    FILENAME_LABEL_H = Inches(0) # Filename text has been removed, so set to 0

    PPTX_TEMPLATE_FILENAME = 'JCET_Format.pptx'

    # 1. Check template file and set path
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
    template_path = os.path.join(script_dir, PPTX_TEMPLATE_FILENAME)
    
    if not os.path.exists(template_path):
        messagebox.showerror("Error", f"PPTX template file not found.\nPath: {template_path}\n(It must be in the same folder as the script)")
        return

    # 2. Set and check the Plots folder path
    try:
        if root_path and os.path.isdir(os.path.join(root_path, "Plots")):
            plots_base_dir = root_path
        else:
            plots_base_dir = filedialog.askdirectory(title="Please select the directory containing the 'Plots' folder.")
            if not plots_base_dir:
                return
    except NameError:
          plots_base_dir = filedialog.askdirectory(title="Please select the directory containing the 'Plots' folder.")
          if not plots_base_dir:
              return

    plots_folder = os.path.join(plots_base_dir, "Plots")
    
    if not os.path.isdir(plots_folder):
        messagebox.showwarning("Warning", f"The 'Plots' folder does not exist in the selected path.\nPlease generate plots first or check the path.")
        return

    log_message.set(f"Generating PPTX report based on {len(SLIDE_PLAN)} defined slides...")
    results.insert(tk.END, "\n\n--- STARTING CUSTOM PPTX GENERATION ---\n", 'subfolder')
    results.yview_moveto(1.0)
        
    total_plots = sum(len(plan['files']) for plan in SLIDE_PLAN)
    if total_plots == 0:
        log_message.set("No images defined in SLIDE_PLAN. PPT generation aborted.")
        results.yview_moveto(1.0)
        return
        
    # 3. Create Presentation object and set slide layout
    try:
        prs = Presentation(template_path)
        SLIDE_LAYOUT_INDEX = 1 
        slide_layout = prs.slide_layouts[SLIDE_LAYOUT_INDEX]
        SLIDE_W = prs.slide_width
        SLIDE_H = prs.slide_height

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load PPTX template: {e}")
        return
    
    # --- Title Slide ---
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Cal Data"
    if len(title_slide.placeholders) > 1:
        title_slide.placeholders[1].text = (
            f"Generated Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Total Slides: {len(SLIDE_PLAN)} (Excluding Title)\n"
            f"Total Plots: {total_plots}"
        )
        
    # 4. Set common placement criteria (Center alignment based on the entire 2x2 block)
    IMG_W = PPTX_IMG_W 
    IMG_H = PPTX_IMG_H
    GAP = PPTX_GAP

    # Width of the entire 2x2 grid (criterion for horizontal center alignment)
    STANDARD_BLOCK_W = (2 * IMG_W) + GAP 
    
    # Height of the entire 2x2 grid (criterion for vertical center alignment)
    TOTAL_2X2_BLOCK_H = (2 * IMG_H) + GAP

    # Maximum usable height area starting below the title
    USABLE_AREA_START_Y = PPTX_TITLE_MARGIN_TOP
    MAX_USABLE_H = SLIDE_H - USABLE_AREA_START_Y - Inches(0.5) 
    
    # Y start position of the entire 2x2 grid block (Vertical center alignment)
    # This aligns to the top even if there are 1 or 2 plots, as it uses the 2x2 block as a reference.
    start_y_for_block = USABLE_AREA_START_Y + (MAX_USABLE_H - TOTAL_2X2_BLOCK_H) / 2
    
    # X start position of the entire block (Horizontal center alignment on the slide)
    block_start_x = (SLIDE_W - STANDARD_BLOCK_W) / 2

    # 5. Generate slides and place images as defined in SLIDE_PLAN
    total_slides = 0
    
    for plan in SLIDE_PLAN:
        
        slide = prs.slides.add_slide(slide_layout)
        current_chunk = plan['files']
        num_plots = len(current_chunk)
        total_slides += 1
        
        slide.shapes.title.text = plan['title']
        
        # --- Image Placement Logic ---
        
        for idx, filename in enumerate(current_chunk):
            
            img_path = os.path.join(plots_folder, filename)
            
            if not os.path.exists(img_path):
                results.insert(tk.END, f"  > WARNING: File not found: {filename}. Skipping.\n", 'not_found')
                continue
                
            row = idx // 2 # 0: Top, 1: Bottom
            col = idx % 2  # 0: Left, 1: Right
            
            # Determine X position (Calculated based on the start point of the 2x2 grid's center block in all cases)
            if col == 0: 
                # Plot 1, 3 (Left column): Place at the starting X position of the entire block.
                left = block_start_x
            else: # col == 1
                # Plot 2, 4 (Right column): Place at the block start + image width + gap position.
                left = block_start_x + IMG_W + GAP

            # Calculate Y position (Using the fixed start_y_for_block)
            top = start_y_for_block + (row * (IMG_H + GAP))
            
            # Insert image
            pic = slide.shapes.add_picture(
                img_path, 
                left, 
                top, 
                width=IMG_W, 
                height=IMG_H
            )

        results.insert(tk.END, f"  > Slide {total_slides} created: '{plan['title']}' with {num_plots} plots.\n", 'saved')
        results.yview_moveto(1.0)

    # 6. Save final file
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    save_filename = f"ppt_{timestamp}.pptx"
    save_path = os.path.join(plots_base_dir, save_filename)
    
    try:
        prs.save(save_path)
        
        msg = f"PPT Generation Complete!\nCustom report saved as: {save_filename}\nTotal Slides: {total_slides} (+ Title)"
        log_message.set(msg)
        messagebox.showinfo("Success", msg)
        results.insert(tk.END, f"\n--- FINAL CUSTOM REPORT SAVED: {save_filename} ---\n", 'subfolder')
        results.yview_moveto(1.0)
        
    except Exception as e:
        results.insert(tk.END, f"  > ERROR saving final PPTX file: {e}\n", 'error')
        messagebox.showerror("Error", f"Failed to save PPTX file:\n{e}")

def quit_application():
    """
    Function to close the tkinter window and exit the Python script.
    """
    root.quit() 
    root.destroy()
    sys.exit() 

# Create the Tkinter window
root = tk.Tk()
root.title("Cal Data Analysis_1.0.1")
root.geometry("600x500") 

# String/Label for status message display
log_message = tk.StringVar()
log_message.set("Click '1. Generate Result Sheet from CSV Folder' to start processing data.")
status_label = tk.Label(root, textvariable=log_message, fg="blue", pady=10)
status_label.pack(fill='x')

# Folder selection button 
select_button = tk.Button(
    root,
    text="1. Generate Result Sheet from CSV Folder",
    command=select_folder_and_process_csv,
    bg="#1E90FF",
    font=("Helvetica", 12, "bold")
)
select_button.pack(pady=(10, 5), padx=20, fill='x')

# Generate Plots button
plot_button = tk.Button(
    root,
    text=f"2. Generate Plots from result sheet", 
    command=generate_plots,
    bg="#32CD32", 
    font=("Helvetica", 12, "bold"),
    state=tk.NORMAL 
)
plot_button.pack(pady=(5, 10), padx=20, fill='x')

# Generate PPT button
plot_button = tk.Button(
    root,
    text=f"3. Generate PPT from plot Folder", 
    command=generate_ppt,
    bg="#FFD700", 
    font=("Helvetica", 12, "bold"),
    state=tk.NORMAL 
)
plot_button.pack(pady=(5, 10), padx=20, fill='x')

# QUIT button
quit_button = tk.Button(
    root,
    text="QUIT",
    command=quit_application,
    bg="#FF6347",
    font=("Helvetica", 12, "bold")
)
quit_button.pack(pady=(5, 15), padx=20, fill='x')

# Frame for results (Text widget and Scrollbar)
results_frame = tk.Frame(root) 
results_frame.pack(pady=10, padx=20, fill='both', expand=True) 

scrollbar = tk.Scrollbar(results_frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Text widget to display processing results
results = tk.Text(
    results_frame,
    wrap=tk.WORD,
    yscrollcommand=scrollbar.set,
    height=15
)
results.pack(side=tk.LEFT, fill='both', expand=True)
scrollbar.config(command=results.yview)

# Apply style tags to the Text widget
results.tag_configure('subfolder', font=('Helvetica', 10, 'bold'), foreground='darkgreen')
results.tag_configure('not_found', font=('Helvetica', 10, 'italic'), foreground='orange')
results.tag_configure('error', font=('Helvetica', 10, 'bold'), foreground='red')
results.tag_configure('saved', font=('Helvetica', 10, 'bold'), foreground='purple')

# GUI Execution
root.mainloop()