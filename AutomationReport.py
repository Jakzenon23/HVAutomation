import pandas as pd
import os

def generate_filtered_unique_assets_from_folder(
    input_folder="DWG_Reports",
    output_folder="Output_Reports",
    filtered_filename="3b_filtered_unique_assets.xlsx",
    z_threshold=81610
):
    """
    Processes all Excel files in the input folder and generates a filtered unique assets report:
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

    excel_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.xlsx', '.xls'))]
    if not excel_files:
        print("No Excel files found in the input folder.")
        return

    filtered_dataframes = []
    for file in excel_files:
        file_path = os.path.join(input_folder, file)
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            df.columns = df.columns.str.strip()  # Strip extra spaces in column names

            # Check if required columns exist
            required_cols = {"Z", "Name", "ID", "Width (mm)", "Depth (mm)", "Height (mm)"}
            missing_cols = required_cols - set(df.columns)
            if missing_cols:
                print(f"Skipping '{file}': Missing columns {missing_cols}")
                continue

            # Filter data
            filtered_df = df[(df["Z"] < z_threshold) & (~df["Name"].isin(bypass_assets))]
            filtered_df = filtered_df[["Name", "ID", "Width (mm)", "Depth (mm)", "Height (mm)"]]

            # Remove duplicates based on 'Name'
            filtered_df = filtered_df.drop_duplicates(subset=["Name"])

            filtered_dataframes.append(filtered_df)

        except Exception as e:
            print(f"Error processing '{file}': {e}")

    if filtered_dataframes:
        combined_filtered_df = pd.concat(filtered_dataframes, ignore_index=True)
        os.makedirs(output_folder, exist_ok=True)
        output_path = os.path.join(output_folder, filtered_filename)
        combined_filtered_df.to_excel(output_path, index=False, engine="openpyxl")
        print(f"Filtered unique assets report saved to: {output_path}")
    else:
        print("No valid data found after filtering.")

