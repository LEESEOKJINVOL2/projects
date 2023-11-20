import os
import pandas as pd
import matplotlib.pyplot as plt
from tkinter import filedialog, Tk, Button
from matplotlib.ticker import StrMethodFormatter

def convert():
    print("convert")

def load_csv():
    file_paths = filedialog.askopenfilenames(title="Select CSV Files", filetypes=[("CSV files", "*.csv")]) # load csv

    if not file_paths:
        print("No file selected. Exiting.")
        return None

    dfs = []
    for file_path in file_paths:
        try:
            df = pd.read_csv(file_path)
            dfs.append((df, file_path))
        except pd.errors.EmptyDataError:
            print(f"Warning: Empty file found - {file_path}")

    if not dfs:
        print("No valid data found. Exiting.")
        return None

    return dfs

def save_png(data_list, save_folder):
    for data, csv_file_path in data_list:
        csv_file_name = os.path.basename(csv_file_path) # csv file name
        fig, ax = plt.subplots(figsize=(10, 6)) # graph size

        psrr_data = data["PSRR"] # csv category: PSRR
        freq_data = data["Frequency"] # csv category: Frequency

        ax.plot(freq_data, psrr_data, linestyle='-') # x: Frequency / y: PSRR
        ax.set_xscale('log') # log scale
        ax.set_xticks([0.01, 0.1, 1, 10, 100]) # show frequency: 0.01 0.1 1 10 100 kHz
        ax.get_xaxis().set_major_formatter(StrMethodFormatter('{x}')) # decimal
        ax.set_xlabel("Frequency (kHz)") # x label: Frequency
        ax.set_ylabel("Gain (dB)") # y lable: Gain
        ax.set_ylim([0, 100]) # y span: 0 to 100
        ax.set_title("PSRR") # graph title
        ax.grid(True) # gird

        png_file_path = os.path.join(save_folder, os.path.splitext(csv_file_name)[0] + ".png") # png file name = csv file name
        plt.savefig(png_file_path)
        plt.close()

def load_save():
    data_list = load_csv()

    if data_list is not None:
        save_folder = filedialog.askdirectory(title="Select Save Folder")
        if save_folder:
            save_png(data_list, save_folder)

def quit():
    root.destroy()

root = Tk()

convert_button = Button(root, text="Convert", command=convert) # 
convert_button.pack()

open_button = Button(root, text="Load & Save", command=load_save) # csv file load & save png file
open_button.pack()

quit_button = Button(root, text="Quit", command=quit) # quit
quit_button.pack()

root.mainloop()