import streamlit as st
import pandas as pd
import os

def combine_excel_files_side_by_side(input_folder="DWG_Reports", 
                                     output_folder="Output_Reports", 
                                     output_filename="combined_report.xlsx"):
    """
    Reads all Excel files in the input folder and places them side by side
    (concatenate by columns). The resulting DataFrame is saved as an Excel file
    in the output folder.
    
    Note: If the DataFrames have different row counts, rows in the smaller
    DataFrame(s) will be filled with NaN to match the largest DataFrame.
    """
    # List all files that end with .xlsx or .xls (case-insensitive)
    excel_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.xlsx', '.xls'))]

    st.write(f"Looking for Excel files in '{input_folder}'...")
    st.write("Excel files found:", excel_files)

    if not excel_files:
        st.write(f"No Excel files found in {input_folder}.")
        return

    dataframes = []
    for file in excel_files:
        file_path = os.path.join(input_folder, file)
        st.write(f"Reading: {file_path}")
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            dataframes.append(df)
        except Exception as e:
            st.write(f"Error reading file {file_path}: {e}")

    # If we successfully read at least one DataFrame
    if dataframes:
        # Concatenate side by side using axis=1
        combined_df = pd.concat(dataframes, axis=1, ignore_index=False)

        # Ensure the output folder exists
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        output_path = os.path.join(output_folder, output_filename)
        combined_df.to_excel(output_path, index=False, engine="openpyxl")
        st.write(f"Combined (side-by-side) Excel file saved to: {output_path}")
    else:
        st.write("No valid Excel files to combine.")
