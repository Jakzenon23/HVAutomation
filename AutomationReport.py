import pandas as pd
import os

def combine_excel_files(input_folder="DWG_Reports", output_folder="Output_Reports", output_filename="combined_report.xlsx"):
    """
    Reads all Excel files in the input folder, concatenates them into one DataFrame,
    and saves the combined DataFrame as an Excel file in the output folder.
    """
    # List all Excel files in the input folder
    excel_files = [os.path.join(input_folder, file) for file in os.listdir(input_folder)
                   if file.endswith(('.xlsx', '.xls'))]
    
    if not excel_files:
        print(f"No Excel files found in {input_folder}.")
        return

    dataframes = []
    for file in excel_files:
        try:
            df = pd.read_excel(file)
            dataframes.append(df)
        except Exception as e:
            print(f"Error reading file {file}: {e}")
    
    if dataframes:
        combined_df = pd.concat(dataframes, ignore_index=True)
        # Ensure output folder exists
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        output_path = os.path.join(output_folder, output_filename)
        combined_df.to_excel(output_path, index=False, engine="openpyxl")
        print(f"Combined Excel file saved to: {output_path}")
    else:
        print("No valid Excel files to combine.")
