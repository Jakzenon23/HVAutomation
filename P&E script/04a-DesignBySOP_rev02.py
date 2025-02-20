import pandas as pd
import os

# Clear the console
clear = lambda: os.system('cls')
clear()

# Define file paths
file1_path = r"C:\Reports\2025\P&E\03b-CombinedSOPsDIMsFilterBaseLevel.xlsx"  # List of assets
file2_path = r"C:\Reports\2025\P&E\04a-DesignFoundationTypes.xlsx"  # Foundation types file
output_file = r"C:\Reports\2025\P&E\04a-DesignFoundationBySOPs.xlsx"

# Load the asset list (File 1)
df_assets = pd.read_excel(file1_path)

# Load the foundation types (File 2)
df_foundations = pd.read_excel(file2_path)

# Merge assets with foundation type, excluding special asset and offsets
df_merged = df_assets.merge(
    df_foundations[['Name', 'Width (mm)', 'Depth (mm)', 'Height (mm)', 'Foundation Type']],
    on=['Name', 'Width (mm)', 'Depth (mm)', 'Height (mm)'],
    how='left'
)

# Now, bring in the special asset information separately
df_foundation_info = df_foundations[['Name', 'Width (mm)', 'Depth (mm)', 'Height (mm)', 
                                     'special asset', 'offset A', 'offset B']]

# Merge again just to get special asset flags and offsets
df_merged = df_merged.merge(df_foundation_info, 
                            on=['Name', 'Width (mm)', 'Depth (mm)', 'Height (mm)'], 
                            how='left')

# Process special assets
new_rows = []
for _, row in df_merged.iterrows():
    row = row.copy()  # Ensure row modification doesn't affect the original DataFrame
    rotation = row["Rotation"]  # Get rotation value
    
    # Determine which coordinate to modify
    coord_to_modify = "Y" if rotation == 90 else "X"
    
    if row['special asset'] == 'yes':
        # Modify the appropriate coordinate in the original row
        row[coord_to_modify] -= row['offset A']
        new_rows.append(row)  # Add modified original row
        
        # Create second row with modified coordinate
        new_row = row.copy()
        new_row[coord_to_modify] += (row['offset A'] + row['offset B'])  # Adjust new coordinate
        new_row.iloc[0] = 'Idem (special asset)'  # Replace first column with "Idem (special asset)"
        new_rows.append(new_row)
    else:
        new_rows.append(row)  # Add normal row

# Convert to DataFrame
df_final = pd.DataFrame(new_rows)

# Drop unnecessary columns
df_final.drop(columns=['special asset', 'offset A', 'offset B','Rotation'], inplace=True)

# Save to Excel
df_final.to_excel(output_file, index=False)

print(f"✅ Process completed! Output saved to {output_file}")
