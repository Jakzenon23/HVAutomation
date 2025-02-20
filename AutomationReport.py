import streamlit as st
import pandas as pd
import os

def combine_excel_files(input_folder="DWG_Reports", 
                        output_folder="Output_Reports", 
                        output_filename="combined_report.xlsx"):
    """
    Reads all Excel files in the input folder, concatenates them,
    and saves the combined DataFrame as an Excel file in the output folder.
    """
    # List all files that end with .xlsx or .xls (case-insensitive)
    excel_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.xlsx', '.xls'))]

    st.write(f"Looking for Excel files in '{input_folder}'...")
    st.write("Excel files found:", excel_files)

    # If no Excel files, exit early
    if not excel_files:
        st.write(f"No Excel files found in {input_folder}.")
        return

    dataframes = []
    # Read each Excel file into a DataFrame
    for file in excel_files:
        file_path = os.path.join(input_folder, file)
        st.write(f"Reading: {file_path}")
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            dataframes.append(df)
        except Exception as e:
            st.write(f"Error reading file {file_path}: {e}")

    # Combine the DataFrames if any were successfully read
    if dataframes:
        combined_df = pd.concat(dataframes, ignore_index=True)

        # Ensure the output folder exists
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        output_path = os.path.join(output_folder, output_filename)
        combined_df.to_excel(output_path, index=False, engine="openpyxl")
        st.write(f"Combined Excel file saved to: {output_path}")
    else:
        st.write("No valid Excel files to combine.")
