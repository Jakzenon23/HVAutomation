import pandas as pd
import os
import stat
import shutil

def change_permissions(path):
    """Change permissions of files in the directory to ensure we can write to them."""
    for foldername, subfolders, filenames in os.walk(path):
        for filename in filenames:
            filepath = os.path.join(foldername, filename)
            os.chmod(filepath, stat.S_IWRITE)  # Remove read-only attribute

def clear_folder(path):
    """Clear the contents of the folder before writing to it."""
    for filename in os.listdir(path):
        file_path = os.path.join(path, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
        except PermissionError:
            print(f"Permission error while trying to delete file: {file_path}")

def factorstemplate(mat):
    factors = open("Templates\\MaterialsBYLevels.csv", "r")
    factors = factors.read().split('\n')
    factors.pop(0)
    for data in factors:
        data = data.replace(" ", "")
        data = data.split(',')
        tempmat = data[0]
        factor = data[2]
        density = data[5]
        if str(tempmat) == str(mat):
            return factor, density
    print('!!! Material not on Factors.csv list !!!', mat)
    input()

def mattemplate(level):
    materials = open("Templates\\Factors.csv", "r")
    materials = materials.read().split('\n')
    materials.pop(0)
    for data in materials:
        data = data.replace(" ", "")
        data = data.split(',')
        templevel = data[0]
        mat = data[1]
        if str(templevel) == str(level):
            return mat
    print('!!! Level not on MaterialsBYLevels.csv list !!!')
    print('Level name: ', level)

def filenames():
    folder_path = "DWG_Reports"
    filenames = []
    # Loop through the files in the folder and append file names to the list
    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            filenames.append(filename)
    sortedlist = sorted(filenames)
    return sortedlist

######################
#### main routine ####
######################
def main(fileRead):
    print("DWG_Reports/" + fileRead)
    f = open("DWG_Reports/" + fileRead, "r")
    file = (f.read()).split('\n')
    file.pop(0)
    
    # Ensure Output_Reports folder is ready
    folder_path = "Output_Reports"
    # Ensure folder permissions and clear any old files
    change_permissions(folder_path)
    clear_folder(folder_path)

    fileWrite = os.path.join(folder_path, fileRead)
    fileout = open(fileWrite, "w")
    fileout.write('Element,Material,Quantity,Carbon Factor (kgCO2e/kg),Carbon (kgCO2e),Density (kg/m3)\n')
    
    dataframe = []
    for line in file:
        if line:
            line = line.replace(" ", "")
            line = line.split(',')
            level = line[1]
            level = level.replace(" ", "")
            volume = line[2]
            # match level in template file
            mat = mattemplate(level)
            # match material in template file
            factor, density = factorstemplate(mat)
            # calculate carbon
            try:
                mass = float(volume) * float(density)
                carbon = round(mass * float(factor), 4)
                data = (level, mat, float(volume), float(factor), float(carbon), float(density))
                dataframe.append(data)
            except:
                print('!!! Error in carbon calculation !!!')
                print('Error data:', mat, volume, density, factor)

            # write to file
            fileout.write(str(level) + ',' + str(mat) + ',' + str(volume) + ',' + str(factor) + ',' + str(carbon) + ',' + str(density))
            fileout.write('\n')

    fileout.close()
    
    # pandas part
    df = pd.DataFrame(dataframe, columns=["Element", "Material", "Quantity", "Carbon Factor", "Carbon", "Density"])
    grouped_df = df.groupby("Element").agg({
        "Element": "first",
        "Material": "first",
        "Quantity": "sum",
        "Carbon Factor": "first",
        "Carbon": "sum",
        "Density": "first"
    })
    
    groupedFileName = "Output_Reports/" + fileRead.replace(".csv", ".xlsx")
    print(groupedFileName)
    grouped_df.to_excel(groupedFileName, index=False, columns=["Element", "Material", "Quantity", "Carbon Factor", "Carbon", "Density"])
    
    print('\n############################', "End file process!", "############################\n")

####################
### Start Script ###
####################
# filenames()
# fileList = filenames()
# for fileRead in fileList:
#    fileRead = str(fileRead)
#    main(fileRead)

# print('All done!!!')
