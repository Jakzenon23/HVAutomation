import streamlit as st
import os
import shutil
import AutomationReport 
from zipfile import ZipFile

def delete_folder(folder_path):
    try:
        shutil.rmtree(folder_path)
    except Exception:
        pass  # Silently ignore or handle errors

# Clean up folders
delete_folder("DWG_Reports")
delete_folder("Output_Reports")

# Create new folders
upload_folder = "DWG_Reports"
os.makedirs(upload_folder, exist_ok=True)

output_folder = "Output_Reports"
os.makedirs(output_folder, exist_ok=True)

st.image("https://www.workspace-interiors.co.uk/application/files/thumbnails/xs/3416/1530/8285/tony_gee_large_logo_no_background.png", width=250)
st.title("Tony Gee Manchester, HV Automation")
st.write("Drag and drop Excel files below to upload them.")

uploaded_files = st.file_uploader("Upload Files", type=["xlsx", "xls"], accept_multiple_files=True)

files_uploaded = False

if uploaded_files:
    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        file_path = os.path.join(upload_folder, file_name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
    files_uploaded = True

if files_uploaded:
    if st.button("Process Uploaded Files"):

        # ✅ Call the updated function that processes directly from the folder
        AutomationReport.generate_filtered_unique_assets_from_folder(
            input_folder=upload_folder,
            output_folder=output_folder,
            filtered_filename="3b_filtered_unique_assets.xlsx"
        )

        filtered_file_path = os.path.join(output_folder, "3b_filtered_unique_assets.xlsx")

        # ✅ Check if the filtered file was created successfully
        if os.path.exists(filtered_file_path):
            st.success("File processing complete!")

            # zip only the filtered report
            zip_path = "processed_files.zip"
            with ZipFile(zip_path, 'w') as zipf:
                zipf.write(filtered_file_path, arcname="3b_filtered_unique_assets.xlsx")

            # provide download button
            with open(zip_path, "rb") as f:
                st.download_button(
                    label="Download Processed Files", 
                    data=f, 
                    file_name="processed_files.zip"
                )
        else:
            st.error("❌ Filtered report could not be generated. Please check if the uploaded files contain the required data.")
