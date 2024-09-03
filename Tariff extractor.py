import tkinter as tk
from tkinter import filedialog
import customtkinter
import pandas as pd
from pyxlsb import open_workbook
from openpyxl import Workbook
import os

# Nastavení pandas pro zobrazení velkého počtu řádků/sloupců
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_colwidth', None)

def selectExcelFile1(pathEntry):
    filename = filedialog.askopenfilename(
        initialdir=os.getcwd(),
        title="Select file",
        filetypes=(("Excel files", "*.xlsx;*;.xlsb"), ("All files", "*.*"))
    )
    pathEntry.delete(0, tk.END)
    pathEntry.insert(0, filename)

def selectExcelFile2(pathEntry):
    filename = filedialog.askopenfilename(
        initialdir=os.getcwd(),
        title="Select file",
        filetypes=(("Excel files", "*.xlsx;*;.xlsb"), ("All files", "*.*"))
    )
    pathEntry.delete(0, tk.END)
    pathEntry.insert(0, filename)

def convert_excel_to_csv(excel_path, csv_path):
    # Zjištění přípony souboru
    file_extension = os.path.splitext(excel_path)[1].lower()

    # Načtení a konverze podle přípony souboru
    if file_extension == '.xlsb':
        print("Načítám .xlsb soubor...")
        df = pd.read_excel(excel_path, sheet_name="FM_Tariffs", engine="pyxlsb")
    elif file_extension in ['.xlsx', '.xlsm']:
        print(f"Načítám {file_extension} soubor...")
        df = pd.read_excel(excel_path, sheet_name="FM_Tariffs", engine="openpyxl")
    else:
        raise ValueError(f"Nepodporovaný typ souboru: {file_extension}")

    # Uložení do CSV
    df.to_csv(csv_path, index=False)
    print(f"Konverze dokončena: {csv_path}")

def extract_origin_destination(od_pair):
    origin = od_pair[:10]
    destination = od_pair[-10:]
    return origin, destination

def load_od_pairs(od_pairs_path):
    # Načtení OD pairů z Excel souboru
    od_df = pd.read_excel(od_pairs_path, engine="openpyxl")
    od_df.columns = ['OD_Pair']  # Přizpůsobte název podle skutečného sloupce
    od_df[['Origin', 'Destination']] = od_df['OD_Pair'].apply(lambda x: pd.Series(extract_origin_destination(x)))
    return od_df

def upload(tariff_path, od_pairs_path, output_excel_path):
    csv_path = "tariffsheet.csv"
    convert_excel_to_csv(tariff_path, csv_path)
    # Load data from CSV
    tariff_Sheet = pd.read_csv(csv_path, header=None, skiprows=1, low_memory=False)

    # Load OD pairs
    od_pairs = load_od_pairs(od_pairs_path)

    # Chargable Weights
    weights = ["200", "600", "1000", "2000", "4000", "10000", "15000", "20000", "25000", ">25000"]

    # Prepare for new format results
    all_results = []

    # Filtering based on OD pairs and creating new data format
    for index, row in od_pairs.iterrows():
        origin = row['Origin']
        destination = row['Destination']
        print(f"Processing OD pair {index + 1}: Origin={origin}, Destination={destination}")

        # Filter based on origin and destination
        filtered_df = tariff_Sheet[(tariff_Sheet.iloc[:, 8].str.contains(origin, na=False) &
                                    tariff_Sheet.iloc[:, 9].str.contains(destination, na=False))]

        if filtered_df.empty:
            print(f"No matching rows found for OD pair {index + 1}")
            continue  # Skip to the next OD pair if no matching rows are found

        # Iterate through filtered rows and create a new format
        for _, filtered_row in filtered_df.iterrows():
            
            c_column_value = str(filtered_row[2])
            d_column_value = str(filtered_row[3])
            r_column_value = filtered_row[17]
            ftl_column_value = filtered_row[19]
            od_pair = f"{origin}__{destination}"

            if "FTL" in d_column_value:
                result = {'OD Pair': od_pair, 'FTL/LTL': "FTL",'Tariff_type': c_column_value, 'equipment': r_column_value, 'Chargable Weight': None, "Cost": ftl_column_value}
                all_results.append(result)
            elif "LTL" in d_column_value:
                # Create a new row for each weight
                i = 40
                for weight in weights:
                    result = {'OD Pair': od_pair, 'FTL/LTL': "LTL", 'Tariff_type': c_column_value, 'equipment': "", 'Chargable Weight': weight, "Cost": filtered_row[i]}
                    i = i + 1
                    all_results.append(result)
                    
        print(f"Finished processing OD pair {index + 1}")

    # Convert results to DataFrame
    final_results = pd.DataFrame(all_results)

    # Format numbers: Ensure 'Cost' and any other numeric columns are formatted correctly
    def format_number(num):
        try:
            num = float(num)
            return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except ValueError:
            return num

    # Apply formatting to the 'Cost' column and any numeric columns if necessary
    final_results['Cost'] = final_results['Cost'].apply(format_number)
    
    # Save results to Excel file
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        final_results.to_excel(writer, sheet_name='Results', index=False)
    


def app():
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")

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
        command=lambda: selectExcelFile2(pathEntry2)
    )
    browseButton2.grid(row=1, column=2, pady=10, padx=5)

    uploadButton = customtkinter.CTkButton(
        frame,
        text="Upload",
        command=lambda: upload(pathEntry2.get(),pathEntry1.get(),"output3.xlsx")
    )
    uploadButton.grid(row=2, column=1, pady=10, padx=5)

    app.mainloop()

app()