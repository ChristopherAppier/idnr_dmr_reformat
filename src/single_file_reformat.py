import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

def main():

    # Prompt the user to select a folder containing the xlsx DMR files
    folder_path = select_folder()

    # Import the xlsx file as a DataFrame and remove the cover tab
    df = import_data(folder_path)

def select_folder():
    # Prompt the user to select a folder using a file dialog
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    folder_path = filedialog.askdirectory(title="Select a folder")

    return folder_path

def import_data(folder_path):
    # Import the xlsx file as a DataFrame and remove the cover tab
    file_path = os.path.join(folder_path, 'data.xlsx')
    df = pd.read_excel(file_path, sheet_name=None)
    df.pop("Cover", None)

    # Printing the names of the remaining tabs
    #print(df.keys())

    return df

if __name__ == "__main__":
    main()