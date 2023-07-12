import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd

def convert_excel_to_csv():
    folder_path = entry.get()

    # Validate the folder path
    if not os.path.isdir(folder_path):
        result_label.config(text="Invalid folder path.", fg="red")
        return

    # Loop through all files in the folder
    for filename in os.listdir(folder_path):
        # Check if file is an Excel file
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            try:
                # Read Excel file into a Pandas dataframe
                df = pd.read_excel(os.path.join(folder_path, filename))
                # Create CSV filename by replacing extension
                csv_filename = os.path.splitext(filename)[0] + ".csv"
                # Write dataframe to CSV file with UTF-8 encoding
                df.to_csv(os.path.join(folder_path, csv_filename), index=False, encoding='utf-8')
                result_label.config(text=f"Converted {filename} to CSV successfully.", fg="green")
            except Exception as e:
                result_label.config(text=f"Error converting {filename} to CSV: {str(e)}", fg="red")
                return

# Create the main window
window = tk.Tk()
window.title("Excel to CSV Converter")

# Create a label and an entry widget for the directory path
path_label = tk.Label(window, text="Enter the folder path:", height=2, width=20)
path_label.pack()
entry = tk.Entry(window, width=50)
entry.pack()

# Create a "Start" button to initiate the conversion process
start_button = tk.Button(window, text="Start", command=convert_excel_to_csv, height=2, width=20)
start_button.pack()

# Create a label to display the conversion result
result_label = tk.Label(window, text="", height=2, width=50)
result_label.pack()

window.mainloop()
