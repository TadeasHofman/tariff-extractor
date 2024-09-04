# -*- coding: utf-8 -*-
# pyinstaller --noconsole --onefile app.py

import tkinter as tk
from tkinter import filedialog
import customtkinter
import pandas as pd
from openpyxl import Workbook
import os
import subprocess
import threading
import time

# Setting pandas options to display a large number of rows/columns
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_colwidth', None)

def selectExcelFile1(pathEntry):
    filename = filedialog.askopenfilename(
        initialdir=os.getcwd(),
        title="Select file",
        filetypes=(("Excel files", "*.xlsx;*.xlsb"), ("All files", "*.*"))
    )
    pathEntry.delete(0, tk.END)
    pathEntry.insert(0, filename)

def selectExcelFile2(pathEntry):
    filename = filedialog.askopenfilename(
        initialdir=os.getcwd(),
        title="Select file",
        filetypes=(("Excel files", "*.xlsx;*.xlsb"), ("All files", "*.*"))
    )
    pathEntry.delete(0, tk.END)
    pathEntry.insert(0, filename)

def open_file_location(file_path):
    """Function to open the file location using the default file explorer."""
    directory = os.path.dirname(file_path)
    if os.name == 'nt':  # If the system is Windows
        os.startfile(directory)
    elif os.name == 'posix':  # If the system is MacOS or Linux
        subprocess.Popen(['open', directory])

def update_status(message):
    """Function to update the status message label."""
    status_label.configure(text=message)
    app.update_idletasks()

def update_progress(progress_bar, percentage_label, progress):
    """Updates the progress bar and percentage label."""
    progress_bar.set(progress)  # Set the progress bar
    percentage_label.configure(text=f"{int(progress * 100)}%")  # Convert to percentage display
    app.update_idletasks()

def simulate_progress(progress_bar, percentage_label, start, end, step=0.01, delay=0.1):
    """Simulates smooth progress bar increments from start to end."""
    current_progress = start  # Start from the given starting point
    while current_progress < end:
        current_progress += step
        current_progress = min(current_progress, end)  # Ensure we don't exceed the end value
        update_progress(progress_bar, percentage_label, current_progress)  # Update progress
        time.sleep(delay)

def convert_excel_to_csv(excel_path, csv_path, progress_var):
    # Determine the file extension and update status with filename
    file_extension = os.path.splitext(excel_path)[1].lower()
    filename = os.path.basename(excel_path)

    update_status(f"Loading {filename}...")
    simulate_progress(progress_bar, percentage_label, 0.0, 0.4)  # Smooth progress from 10% to 40%

    if file_extension == '.xlsb':
        df = pd.read_excel(excel_path, sheet_name="FM_Tariffs", engine="pyxlsb")
    elif file_extension in ['.xlsx', '.xlsm']:
        df = pd.read_excel(excel_path, sheet_name="FM_Tariffs", engine="openpyxl")
    else:
        raise ValueError(f"Unsupported file type: {file_extension}")

    # Save to CSV
    df.to_csv(csv_path, index=False)
    update_status(f"Conversion complete: {csv_path}")

    simulate_progress(progress_bar, percentage_label, 0.4, 0.5)  # Smooth progress from 40% to 50%

def load_od_pairs(od_pairs_path, progress_var):
    # Load OD pairs from Excel file
    filename = os.path.basename(od_pairs_path)
    update_status(f"Loading OD pairs from {filename}...")
    simulate_progress(progress_bar, percentage_label, 0.5, 0.6)  # Smooth progress from 50% to 60%

    od_df = pd.read_excel(od_pairs_path, engine="openpyxl")
    od_df.columns = ['OD_Pair']  # Adjust column name according to actual data
    od_df[['Origin', 'Destination']] = od_df['OD_Pair'].apply(lambda x: pd.Series(extract_origin_destination(x)))

    simulate_progress(progress_bar, percentage_label, 0.6, 0.7)  # Smooth progress from 60% to 70%
    
    return od_df

def extract_origin_destination(od_pair):
    origin = od_pair[:10]
    destination = od_pair[-10:]
    return origin, destination

def run_upload_thread(tariff_path, od_pairs_path, output_excel_path, progress_var):
    # Show progress bar and frame only when upload starts
    progress_bar.grid(row=2, column=0, columnspan=2, pady=10)
    percentage_label.grid(row=2, column=2, padx=5)
    status_label.grid(row=3, column=0, columnspan=3, pady=5)
    
    # Run the upload function in a separate thread to keep the GUI responsive
    thread = threading.Thread(target=upload, args=(tariff_path, od_pairs_path, output_excel_path, progress_var))
    thread.start()

def upload(tariff_path, od_pairs_path, output_excel_path, progress_var):
    csv_path = "tariffsheet.csv"
    
    # Update status and simulate progress for initial loading
    update_status(f"Loading {os.path.basename(tariff_path)}...")
    
    # Load and convert Excel to CSV only once
    convert_excel_to_csv(tariff_path, csv_path, progress_var)

    # Update status for loading CSV data and simulate progress
    update_status("Loading tariff data from CSV...")
    tariff_Sheet = pd.read_csv(csv_path, header=None, skiprows=1, low_memory=False)
    
    update_progress(progress_bar, percentage_label, 0.5)  # Update progress to 50%

    # Load OD pairs only once and update progress
    od_pairs = load_od_pairs(od_pairs_path, progress_var)

    # Process OD pairs and update progress incrementally
    weights = ["200", "600", "1000", "2000", "4000", "10000", "15000", "20000", "25000", ">25000"]
    all_results = []

    for index, row in od_pairs.iterrows():
        origin = row['Origin']
        destination = row['Destination']
        update_status(f"Processing OD pair {index + 1}: Origin={origin}, Destination={destination}")

        filtered_df = tariff_Sheet[
            (tariff_Sheet.iloc[:, 8].str.contains(origin, na=False) & 
             tariff_Sheet.iloc[:, 9].str.contains(destination, na=False))
        ]

        if filtered_df.empty:
            update_status(f"No matching rows found for OD pair {index + 1}")
            continue

        for _, filtered_row in filtered_df.iterrows():
            c_column_value = str(filtered_row[2])
            d_column_value = str(filtered_row[3])
            r_column_value = filtered_row[17]
            ftl_column_value = filtered_row[19]
            od_pair = f"{origin}__{destination}"

            if "FTL" in d_column_value:
                result = {'OD Pair': od_pair, 'FTL/LTL': "FTL", 'Tariff_type': c_column_value, 'equipment': r_column_value, 'Chargable Weight': None, "Cost": ftl_column_value}
                all_results.append(result)
            elif "LTL" in d_column_value:
                i = 40
                for weight in weights:
                    result = {'OD Pair': od_pair, 'FTL/LTL': "LTL", 'Tariff_type': c_column_value, 'equipment': "", 'Chargable Weight': weight, "Cost": filtered_row[i]}
                    i = i + 1
                    all_results.append(result)

        # Incremental progress updates for OD pair processing
        current_progress = 0.7 + (index + 1) * 0.3 / len(od_pairs)
        update_progress(progress_bar, percentage_label, current_progress)
        update_status(f"Finished processing OD pair {index + 1}")

    # Compile and save results
    update_status("Compiling final results...")
    final_results = pd.DataFrame(all_results)
    final_results['Numeric Chargable Weight'] = final_results['Chargable Weight'].replace('>25000', '99999')
    final_results['Numeric Chargable Weight'] = pd.to_numeric(final_results['Numeric Chargable Weight'], errors='coerce')
    final_results['Cost'] = pd.to_numeric(final_results['Cost'], errors='coerce')

    column_order = final_results.columns.tolist()
    chargable_weight_index = column_order.index('Chargable Weight')
    column_order.insert(chargable_weight_index + 1, column_order.pop(column_order.index('Numeric Chargable Weight')))
    final_results = final_results[column_order]

    simulate_progress(progress_bar, percentage_label, 0.9, 1.0)  # Smooth progress from 90% to 100%

    update_status("Saving results to Excel...")
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        final_results.to_excel(writer, sheet_name='Results', index=False)

    update_progress(progress_bar, percentage_label, 1.0)  # Set progress to 100%
    update_status("Processing complete!")

    open_location_button.configure(state=tk.NORMAL)  # Enable the button to open the file location

def app():
    customtkinter.set_appearance_mode("System")  # Set the appearance mode to system (light/dark based on OS settings)
    customtkinter.set_default_color_theme("blue")  # Default color theme

    global app  # Declare app as global to use within functions
    app = customtkinter.CTk()
    app.title("Tariff extractor")

    frame = customtkinter.CTkFrame(app)
    frame.grid(row=0, column=0, sticky="ew", padx=20, pady=20)

    title1 = customtkinter.CTkLabel(frame, text="OD pairs file")
    title1.grid(row=0, column=0, pady=10, padx=5)

    pathEntry1 = customtkinter.CTkEntry(frame)
    pathEntry1.grid(row=0, column=1, pady=10, padx=5)

    browseButton1 = customtkinter.CTkButton(
        frame,
        text="Browse files",
        fg_color="#EE7203",  # Button color
        text_color="#FFFF00",  # Text color inside button
        command=lambda: selectExcelFile1(pathEntry1)
    )
    browseButton1.grid(row=0, column=2, pady=10, padx=5)

    title2 = customtkinter.CTkLabel(frame, text="Tariff file")
    title2.grid(row=1, column=0, pady=10, padx=5)

    pathEntry2 = customtkinter.CTkEntry(frame)
    pathEntry2.grid(row=1, column=1, pady=10, padx=5)

    browseButton2 = customtkinter.CTkButton(
        frame,
        text="Browse files",
        fg_color="#EE7203",  # Button color
        text_color="#FFFF00",  # Text color inside button
        command=lambda: selectExcelFile2(pathEntry2)
    )
    browseButton2.grid(row=1, column=2, pady=10, padx=5)

    # Progress bar and status elements - initially hidden
    global progress_var, progress_bar, percentage_label, status_label, open_location_button
    progress_var = tk.DoubleVar()

    # Use customtkinter's CTkProgressBar for better control over styling
    progress_bar = customtkinter.CTkProgressBar(
        frame,
        orientation='horizontal',
        width=300,
        height=20,
        progress_color="#89CFF0",  # Progress color
        border_color="#EE7203"  # Border color matching the button color
    )
    progress_bar.set(0)  # Set initial progress to 0

    percentage_label = customtkinter.CTkLabel(frame, text="0%")
    status_label = customtkinter.CTkLabel(frame, text="")

    # Upload button
    uploadButton = customtkinter.CTkButton(
        frame,
        text="Upload",
        fg_color="#EE7203",  # Button color
        text_color="#FFFF00",  # Text color inside button
        command=lambda: run_upload_thread(pathEntry2.get(), pathEntry1.get(), "output3.xlsx", progress_var)
    )
    uploadButton.grid(row=4, column=1, pady=10, padx=5)

    # Button to open file location
    open_location_button = customtkinter.CTkButton(
        frame,
        text="Open File Location",
        fg_color="#EE7203",  # Button color
        text_color="#FFFF00",  # Text color inside button
        command=lambda: open_file_location("output3.xlsx"),
        state=tk.DISABLED  # Initially disabled until processing is complete
    )
    open_location_button.grid(row=5, column=1, pady=10, padx=5)

    app.mainloop()

app()
