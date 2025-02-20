import pandas as pd
import os

# Clear console
clear = lambda: os.system('cls')
clear()

bypass_assets = [
    "275 TO 400 SGT HYOSUNG - 275 TO 400 SGT HYOSUNG",
    "Foundation - GantryFoundation",
    "shunt reactor - shunt reactor",
    "SGT 400-132kV - SGT 400-132kV"
]

# Path to the Excel file
file_path = r"C:\Users\JZS\OneDrive - Tony Gee and Partners LLP\Desktop\P&E script\02-ModelQuantitiesDIMs.xlsx"

# Read the Excel file, skipping the first row
df = pd.read_excel(file_path, skiprows=1)

# Convert DataFrame to list of lists (each sublist is a row)
all_rows = df.values.tolist()

# Filter rows where the value in the 4th column (index 3) is less than 81610
filtered_rows_F1 = [row for row in all_rows if row[3] < 81610]
filtered_rows = [row for row in filtered_rows_F1 if row[0] not in bypass_assets]

# Print or further process 'filtered_rows'
print('\n###########################################################################')
print(f"Number of rows where the Z level coord is < 81600: {len(filtered_rows)}")
print('###########################################################################')

# Select columns [0, 1, 2, 3, 4, 7, 8, 9] for output
selected_columns = []
for row in filtered_rows:
    selected_columns.append([row[i] for i in [0, 1, 2, 3, 4, 7, 8, 9, 10]])

# Define the output Excel file path for filtered rows
output_file_path = r"C:\Users\JZS\OneDrive - Tony Gee and Partners LLP\Desktop\P&E script\02-ModelQuantitiesDIMs.xlsx"

# Define the header
header = ["Name", "X", "Y", "Z", "ID", "Width (mm)", "Depth (mm)", "Height (mm)", "Rotation"]

# Create a DataFrame from the selected columns and header
filtered_df = pd.DataFrame(selected_columns, columns=header)

# Save the filtered DataFrame to Excel
filtered_df.to_excel(output_file_path, index=False)

# Output the file path
print(f"Filtered rows saved to {output_file_path}")

# Now create the unique assets file based on 'Name' (row[0]), 'Width' (row[7]), 'Depth' (row[8]), and 'Height' (row[9])
# Create a DataFrame from the original rows
df_assets = pd.DataFrame(filtered_rows, columns=df.columns)

# Remove duplicates based on the combination of 'Name', 'Width', 'Depth', and 'Height'
unique_assets = df_assets.drop_duplicates(subset=[df.columns[0], df.columns[7], df.columns[8], df.columns[9]])

# Select only 'Name', 'Width', 'Depth', and 'Height' for the unique assets report
unique_assets_filtered = unique_assets[[df.columns[0], df.columns[4], df.columns[7], df.columns[8], df.columns[9]]]

# Define the output Excel file path for unique assets
unique_assets_file_path = r"C:\Reports\2025\P&E\03b-CombinedSOPsDIMsUniqueAssetsList.xlsx"

# Save the unique assets DataFrame to Excel
unique_assets_filtered.to_excel(unique_assets_file_path, index=False, header=["Name", "ID", "Width (mm)", "Depth (mm)", "Height (mm)"])

# Output the file path for the unique assets report
print(f"Unique assets saved to {unique_assets_file_path}")