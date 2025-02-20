import streamlit as st
import os
import shutil
import AutomationReport 
from zipfile import ZipFile

def delete_folder(folder_path):
    try:
        shutil.rmtree(folder_path)  # Remove the folder and all its contents (files and subdirectories)
        print(f'Folder {folder_path} deleted successfully.')
    except Exception as e:
        print(f'Failed to delete {folder_path}. Reason: {e}')

# 1. Clean up old folders
delete_folder("DWG_Reports")
delete_folder("Output_Reports")

# 2. Create new folders if they don't exist
upload_folder = "DWG_Reports"
if not os.path.exists(upload_folder):
    os.makedirs(upload_folder)

output_folder = "Output_Reports"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 3. Streamlit UI
st.image("https://www.workspace-interiors.co.uk/application/files/thumbnails/xs/3416/1530/8285/tony_gee_large_logo_no_background.png", width=250)
st.title("Tony Gee Manchester, HV Automation")
st.write("Drag and drop Excel files below to upload them.")

# 4. File uploader widget (allows multiple files)
uploaded_files = st.file_uploader("Upload Files", type=["xlsx", "xls"], accept_multiple_files=True)

# Track whether files were uploaded
files_uploaded = False

# 5. Save each uploaded file to DWG_Reports
if uploaded_files:
    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        file_path = os.path.join(upload_folder, file_name)

        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.success(f"File '{file_name}' uploaded successfully!")

    # Set flag to True if files have been uploaded
    files_uploaded = True

# 6. If files have been uploaded, show a button to process them
if files_uploaded:
    if st.button("Process Uploaded Files"):
        st.write("Processing files...")

        # Debug: list the files in DWG_Reports
        dwg_contents = os.listdir(upload_folder)
        st.write("Files in DWG_Reports folder:", dwg_contents)

        # Call the new combine function from AutomationReport
        AutomationReport.combine_excel_files(
            input_folder=upload_folder,
            output_folder=output_folder,
            output_filename="combined_report.xlsx"
        )

        # Debug: list the files in Output_Reports
        output_contents = os.listdir(output_folder)
        st.write("Files in Output_Reports folder:", output_contents)

        st.success("File processing complete!")

        # 7. Zip everything in Output_Reports and provide download
        zip_path = "processed_files.zip"
        with ZipFile(zip_path, 'w') as zipf:
            for root, dirs, files in os.walk(output_folder):
                for file in files:
                    zipf.write(os.path.join(root, file), arcname=file)

        # Download button for the processed zip file
        with open(zip_path, "rb") as f:
            st.download_button(
                label="Download Processed Files", 
                data=f, 
                file_name="processed_files.zip"
            )
