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
    def generate_filtered_unique_assets(
    combined_file_path,
    output_folder="Output_Reports",
    filtered_filename="3b_filtered_unique_assets.xlsx",
    z_threshold=81610
):
    """
    Generates a filtered unique assets report from a combined Excel file:
    - Rows where Z < z_threshold
    - Excludes bypass assets
    - Unique assets based on 'Name' only
    - Output saved as 3b_filtered_unique_assets.xlsx
    """
    bypass_assets = [
        "275 TO 400 SGT HYOSUNG - 275 TO 400 SGT HYOSUNG",
        "Foundation - GantryFoundation",
        "shunt reactor - shunt reactor",
        "SGT 400-132kV - SGT 400-132kV"
    ]

    # ✅ Read the combined Excel file
    df = pd.read_excel(combined_file_path, engine="openpyxl")
    df.columns = df.columns.str.strip()  # Clean column names

    # ✅ Filter rows based on Z threshold and bypass assets
    filtered_df = df[(df["Z"] < z_threshold) & (~df["Name"].isin(bypass_assets))]

    # ✅ Select required columns
    filtered_df = filtered_df[["Name", "ID", "Width (mm)", "Depth (mm)", "Height (mm)"]]

    # ✅ Remove duplicates based on 'Name'
    unique_filtered_df = filtered_df.drop_duplicates(subset=["Name"])

    # ✅ Save the filtered report
    os.makedirs(output_folder, exist_ok=True)
    filtered_output_path = os.path.join(output_folder, filtered_filename)
    unique_filtered_df.to_excel(filtered_output_path, index=False, engine="openpyxl")

    print(f"🎯 Filtered unique assets report saved to: {filtered_output_path}")
