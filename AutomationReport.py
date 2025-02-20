import pandas as pd
import io

def combine_uploaded_excel_files(uploaded_files, output_filename="combined_report.xlsx"):
    """
    Reads multiple Excel files from a list of uploaded file-like objects,
    concatenates them into one DataFrame, and returns a BytesIO object
    containing the combined Excel file.
    """
    dataframes = []
    for file in uploaded_files:
        try:
            # Read each file using Pandas (assumes Excel file)
            df = pd.read_excel(file)
            dataframes.append(df)
        except Exception as e:
            print(f"Error reading file {file.name}: {e}")

    if dataframes:
        combined_df = pd.concat(dataframes, ignore_index=True)
        output = io.BytesIO()
        combined_df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)
        return output
    else:
        return None
