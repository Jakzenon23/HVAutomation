import streamlit as st
import os
import shutil
import AutomationReport 
import time
from zipfile import ZipFile

def delete_folder(folder_path):
    try:
        shutil.rmtree(folder_path)  # Remove the folder and all its contents (files and subdirectories)
        print(f'Folder {folder_path} deleted successfully.')
    except Exception as e:
        print(f'Failed to delete {folder_path}. Reason: {e}')

# Delete folders
delete_folder("DWG_Reports")
delete_folder("Output_Reports")

# Folders where the uploaded files will be saved and where output will be stored
upload_folder = "DWG_Reports"
if not os.path.exists(upload_folder):
    os.makedirs(upload_folder)
output_folder = "Output_Reports"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Streamlit title and description
st.image("https://www.workspace-interiors.co.uk/application/files/thumbnails/xs/3416/1530/8285/tony_gee_large_logo_no_background.png", width=250)
st.title("Tony Gee Manchester, HV Automation")
st.write("Drag and drop Excel files below to upload them.")

# File uploader widget (allows multiple files)
uploaded_files = st.file_uploader("Upload Files", type=["xlsx", "xls"], accept_multiple_files=True)

# Variable to track if files were uploaded
files_uploaded = False

# Process and save each uploaded file
if uploaded_files:
    for uploaded_file in uploaded_files:
        # Get the file details
        file_name = uploaded_file.name
        file_path = os.path.join(upload_folder, file_name)

        # Save the uploaded file to the folder
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Show a success message for each uploaded file
        st.success(f"File '{file_name}' uploaded successfully!")

    # Set flag to true if files have been uploaded
    files_uploaded = True

# Display a button to run code if files have been uploaded
if files_uploaded:
    if st.button("Process Uploaded Files"):
        st.write("Processing files...")

        # Instead of processing each file individually, call the new combine function
        AutomationReport.combine_excel_files(
            input_folder=upload_folder,
            output_folder=output_folder,
            output_filename="combined_report.xlsx"
        )

        st.success("Files combined successfully!")

        # Zip the output folder for download (will include the combined report)
        zip_path = "processed_files.zip"
        with ZipFile(zip_path, 'w') as zipf:
            for root, dirs, files in os.walk(output_folder):
                for file in files:
                    zipf.write(os.path.join(root, file), arcname=file)

        # Download Button for Processed Files
        with open(zip_path, "rb") as f:
            st.download_button(
                label="Download Processed Files", 
                data=f, 
                file_name="processed_files.zip"
            )
