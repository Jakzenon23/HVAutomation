import pandas as pd
import os

def combine_excel_files_side_by_side(
    input_folder="DWG_Reports", 
    output_folder="Output_Reports", 
    output_filename="combined_report.xlsx"
):
    """
    Reads all Excel files in the input folder and places them side by side
    (concatenates columns). The resulting DataFrame is saved as an Excel file
    in the output folder.
    """
    excel_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.xlsx', '.xls'))]
    if not excel_files:
        return  # No Excel files found; do nothing

    dataframes = []
    for file in excel_files:
        file_path = os.path.join(input_folder, file)
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            # If this is the specific file, drop the specified columns
            if file == "02-ModelQuantitiesSOPS.xlsx":
                df = df.drop(columns=["Name", "ID"], errors="ignore")
            dataframes.append(df)
        except:
            pass  # Silently ignore errors or handle them as you prefer

    if dataframes:
        combined_df = pd.concat(dataframes, axis=1, ignore_index=False)
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        output_path = os.path.join(output_folder, output_filename)
        combined_df.to_excel(output_path, index=False, engine="openpyxl")

def combine_excel_files_foundation_level(
    input_folder="DWG_Reports", 
    output_folder="Output_Reports", 
    output_filename="combined_reportFoundationLevel.xlsx",
    filter_column="Level",
    filter_value="Foundation"
):
    """
    Reads all Excel files in the input folder and concatenates them side by side.
    Then, filters the combined DataFrame to include only rows where filter_column equals filter_value.
    The resulting DataFrame is saved as an Excel file in the output folder.
    """
    excel_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.xlsx', '.xls'))]
    if not excel_files:
        return  # No Excel files found; do nothing

    dataframes = []
    for file in excel_files:
        file_path = os.path.join(input_folder, file)
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            # Drop columns for the specific file if present
            if file == "02-ModelQuantitiesSOPS.xlsx":
                df = df.drop(columns=["Name", "ID"], errors="ignore")
            dataframes.append(df)
        except:
            pass  # Silently ignore errors or handle them as you prefer

    if dataframes:
        combined_df = pd.concat(dataframes, axis=1, ignore_index=False)

        # Filter the DataFrame
        if filter_column in combined_df.columns:
            filtered_df = combined_df[combined_df[filter_column] == filter_value]
        else:
            filtered_df = combined_df  # If the column doesn't exist, no filtering is applied

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        output_path = os.path.join(output_folder, output_filename)
        filtered_df.to_excel(output_path, index=False, engine="openpyxl")




# add a function for foundation level assets
def generate_filtered_unique_assets(
    combined_file_path,
    output_folder="Output_Reports",
    filtered_filename="filtered_unique_assets.xlsx",
    z_threshold=81610
):
    """
    Generates a single filtered unique assets report based on:
    - Z < z_threshold
    - Excludes bypass assets
    - Shows only unique assets based on Name, Width, Depth, and Height
    """
    # * Assets to bypass
    bypass_assets = [
        "275 TO 400 SGT HYOSUNG - 275 TO 400 SGT HYOSUNG",
        "Foundation - GantryFoundation",
        "shunt reactor - shunt reactor",
        "SGT 400-132kV - SGT 400-132kV"
    ]

    # * Read the combined Excel file
    df = pd.read_excel(combined_file_path, engine="openpyxl")

    # * Filter rows where Z < z_threshold and exclude bypassed assets
    filtered_df = df[(df["Z"] < z_threshold) & (~df["Name"].isin(bypass_assets))]

    # * Select required columns for the final output
    filtered_df = filtered_df[["Name", "X", "Y", "Z", "ID", "Width (mm)", "Depth (mm)", "Height (mm)", "Rotation"]]

    # * Remove duplicates based on Name, Width, Depth, and Height (unique assets)
    unique_filtered_df = filtered_df.drop_duplicates(subset=["Name", "Width (mm)", "Depth (mm)", "Height (mm)"])

    # * Save the final filtered unique assets report
    os.makedirs(output_folder, exist_ok=True)
    filtered_output_path = os.path.join(output_folder, filtered_filename)
    unique_filtered_df.to_excel(filtered_output_path, index=False)


