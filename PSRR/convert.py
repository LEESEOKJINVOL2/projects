from tkinter import filedialog, Tk, Button
import pandas as pd
import os

def open_file_dialog():
    file_paths = filedialog.askopenfilenames(title="Select CSV files", filetypes=[("CSV files", "*.csv")])
    return file_paths

def save_folder_dialog():
    folder_path = filedialog.askdirectory(title="Select Folder")
    return folder_path

def convert():
    selected_files = open_file_dialog()
    save_folder = save_folder_dialog()

    if not selected_files:
        print("No files selected.")
        return

    for file_path in selected_files:
        df = pd.read_csv(file_path, skiprows=3)

        combined_df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

        combined_df["       Frequency"] = combined_df["       Frequency"].str.replace(" kHz", "").str.replace(" Hz", "")
        print("1st convert:", combined_df["       Frequency"])
        df.iloc[:, 50:] = df.iloc[:, 50:].astype(float) / 1000
        print("2nd convert:", combined_df["       Frequency"])


        combined_df["            PSRR"] = combined_df["            PSRR"].str.replace(" dB", "")

        if save_folder:
            # 원본 파일 이름과 동일한 이름으로 저장
            file_name = os.path.basename(file_path)
            save_path = os.path.join(save_folder, file_name)
            combined_df.to_csv(save_path, index=False)
            print(f"Processed data saved to '{save_path}'.")

def quit():
    root.destroy()

root = Tk()

convert_button = Button(root, text="Convert", command=convert)
convert_button.pack()

quit_button = Button(root, text="Quit", command=quit)
quit_button.pack()

root.mainloop()