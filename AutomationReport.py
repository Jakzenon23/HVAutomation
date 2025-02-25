import pandas as pd
import os

def combine_excel_files_side_by_side(
    input_folder="DWG_Reports", 
    output_folder="Output_Reports", 
    output_filename="3a_combined_report.xlsx"
):
    """
    Reads all Excel files in the input folder and places them side by side (concatenates columns).
    The resulting DataFrame is saved as an Excel file in the output folder.
    """
    excel_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.xlsx', '.xls'))]
    if not excel_files:
        print("No Excel files found in the input folder.")
        return

    dataframes = []
    for file in excel_files:
        file_path = os.path.join(input_folder, file)
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            # If this is the specific file, drop the specified columns
            if file == "02-ModelQuantitiesSOPS.xlsx":
                df = df.drop(columns=["Name", "ID"], errors="ignore")
            dataframes.append(df)
        except Exception as e:
            print(f"Error processing {file}: {e}")

    if dataframes:
        combined_df = pd.concat(dataframes, axis=1, ignore_index=False)
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        output_path = os.path.join(output_folder, output_filename)
        combined_df.to_excel(output_path, index=False, engine="openpyxl")
        print(f"Combined report saved to: {output_path}")
