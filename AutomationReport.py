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




# add a function for foundation level asset

def generate_filtered_unique_assets(
    combined_file_path,
    output_folder="Output_Reports",
    filtered_filename="filtered_unique_assets.xlsx",
    z_threshold=81610
):
    """
    Generates a filtered unique assets report:
    - Filters rows where Z < z_threshold
    - Excludes bypass assets
    - Selects exact columns matching the original output
    - Deduplicates based on 'Name', 'Width (mm)', 'Depth (mm)', 'Height (mm)'
    - Outputs a single Excel file with: Name, ID, Width (mm), Depth (mm), Height (mm)
    """

    # * Assets to bypass
    bypass_assets = [
        "275 TO 400 SGT HYOSUNG - 275 TO 400 SGT HYOSUNG",
        "Foundation - GantryFoundation",
        "shunt reactor - shunt reactor",
        "SGT 400-132kV - SGT 400-132kV"
    ]

    # * Read the combined Excel file, skipping the first row (matches original code)
    df = pd.read_excel(combined_file_path, engine="openpyxl", skiprows=1)

    # * Convert DataFrame to list of lists (replicating original logic)
    all_rows = df.values.tolist()

    # * Step 1: Filter rows where Z (4th column, index 3) < z_threshold
    filtered_rows_F1 = [row for row in all_rows if row[3] < z_threshold]

    # * Step 2: Exclude bypass assets
    filtered_rows = [row for row in filtered_rows_F1 if row[0] not in bypass_assets]

    # * Step 3: Convert filtered rows back to DataFrame
    filtered_df = pd.DataFrame(filtered_rows, columns=df.columns)

    # * Step 4: Deduplicate based on original columns: Name, Width, Depth, Height
    unique_assets = filtered_df.drop_duplicates(subset=[df.columns[0], df.columns[7], df.columns[8], df.columns[9]])

    # * Step 5: Select only required columns for the final output
    unique_assets_filtered = unique_assets[[df.columns[0], df.columns[4], df.columns[7], df.columns[8], df.columns[9]]]
    unique_assets_filtered.columns = ["Name", "ID", "Width (mm)", "Depth (mm)", "Height (mm)"]

    # * Step 6: Save the final filtered unique assets report
    os.makedirs(output_folder, exist_ok=True)
    filtered_output_path = os.path.join(output_folder, filtered_filename)
    unique_assets_filtered.to_excel(filtered_output_path, index=False)

