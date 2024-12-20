#File Manager

#LIBRARIES
import pandas as pd
import shutil
import os

#LOAD LOG
log = r"C:\Users\nathan.stephens\OneDrive - Sussex County Government\Files\NKS_2024 P&D Log.xlsx"

#REFERENCE FOLDERS
deed_folder = r"C:\Users\nathan.stephens\OneDrive - Sussex County Government\Files\Deeds"
plot_folder = r"C:\Users\nathan.stephens\OneDrive - Sussex County Government\Files\Plots"
wr_folder = r"C:\Users\nathan.stephens\OneDrive - Sussex County Government\Files\Work Requests"

#REFERENCE DOCUMENTS
sd_file = r"C:\Users\nathan.stephens\OneDrive - Sussex County Government\Files\Sub Divisions.xlsx"

#READ LOG, REF BY PAGE
plot_log = pd.read_excel(log, sheet_name = "PLOT", header = 7)
deed_log = pd.read_excel(log, sheet_name = "DEED", header = 7)
wr_log = pd.read_excel(log, sheet_name = "W_R", header = 4)

#SET SHEET PKs
plot_pk = "Plot Book & Page"
deed_pk = "Deed Book & Page"
wr_pk = "ID"

#CREATE FOLDER, MOVE FILE FUNCTION
def create_folder_move_file(folder_path, file_name, destination_folder):
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    file_path = os.path.join(folder_path, file_name)
    if os.path.exists(file_path):
        shutil.move(file_path, destination_folder)

#START COUNTERS
sd_files_copied = 0
files_moved = 0
folders_created = set()
wr_files_moved = 0

#PROCESS PLOT_LOG
for _, row in plot_log.iterrows():
    if row["Status"] == "Compelte":
        continue

    #CREATE FOLDERS
    if row["Work Request"] == "NOT DONE":
        plot_file_name = str(row[plot_pk]) + ".pdf"
        plot_folder_path = os.path.join(plot_folder, str(row[plot_pk]))
        
        #CREATE FOLDER AND MOVE FILE
        create_folder_move_file(plot_folder, plot_file_name, plot_folder_path)
        folders_created.add(str(row[plot_pk]))

        #COPY SD
        if row["Work Request"] == "NOT DONE":
            sd_dest = os.path.join(plot_folder_path, str(row[plot_pk]) + "_SubDiv.xlsx")
            shutil.copy(sd_file, sd_dest)
            sd_files_copied += 1

    #MATCH PLOT AND DEED PKs
    for _, deed_row in deed_log.iterrows():
        if row[plot_pk] == deed_row[deed_pk]:
            deed_file_name = str(deed_row[deed_pk]) + ".docx"
            deed_folder_path = os.path.join(deed_folder, str(deed_row[deed_pk]))

            #MOVE DEED FILE TO PLOT FOLDER
            create_folder_move_file(deed_folder, deed_file_name, plot_folder_path)
            files_moved += 1

    #MATCH WR FILE FROM LOG
    for _, wr_row in wr_log.iterrows():
        if wr_row[wr_pk] == row[plot_pk] and wr_row["Status"] == "COMPLETE":
            wr_file_name = str(wr_row[wr_pk]) + ".docx"
            create_folder_move_file(wr_folder, wr_file_name, plot_folder_path)
            wr_files_moved += 1

#PROCESS DEED_LOG
for _, row in deed_log.iterrows():
    if row["Status"] == "Complete":
        continue
        
    #CREATE FOLDERS
    if row["Work Request"] == "NOT DONE":
        deed_file_name = str(row[deed_pk]) + ".pdf"
        deed_folder_path = os.path.join(deed_folder, str(row[deed_pk]))

        #CREATE FOLDER AND MOVE FILE
        create_folder_move_file(deed_folder, deed_file_name, deed_folder_path)
        folders_created.add(str(row[deed_pk]))

    #MATCH DEED AND PLOT PKs
    for _, plot_row in plot_log.iterrows():
        if row[deed_pk] == plot_row[plot_pk]:
            plot_file_name = str(plot_row[plot_pk]) + ".pdf"
            plot_folder_path = os.path.join(plot_folder, str(plot_row[plot_pk]))
                                            
            #MOVE PLOT FILE TO DEED FOLDER LOCATION
            create_folder_move_file(plot_folder, plot_file_name, deed_folder_path)
            files_moved += 1

    #MATCH WR FILE FROM LOG
    for _, wr_row in wr_log.iterrows():
        if wr_row[wr_pk] == row[deed_pk] and wr_row["Status"] == "COMPLETE":
            wr_file_name = str(wr_row[wr_pk]) + ".docx"
            create_folder_move_file(wr_folder, wr_file_name, deed_folder_path)
            wr_files_moved += 1

#CLEAN UP: MOVE FILES BASED ON STATUS
def move_to_status_folder(folder_path, file_name, status):
    if status == "Complete":
        subfolder = file_name[:3]
        status_folder = os.path.join(folder_path, subfolder, "Complete")
        if not os.path.exists(status_folder):
            os.makedirs(status_folder)
        shutil.move(os.path.join(folder_path, file_name), status_folder)
    else:
        todo_folder = os.path.join(folder_path, "To Do")
        if not os.path.exists(todo_folder):
            os.makedirs(todo_folder)
        shutil.move(os.path.join(folder_path, file_name), todo_folder)

#CLEAN UP: PLOT FOLDER
for _, row in plot_log.iterrows():
    if row["Status"] == "Complete":
        plot_file_name = str(row[plot_pk]) + ".pdf"
        move_to_status_folder(plot_folder, plot_file_name, row["Status"])

#CLEAN UP: DEED FOLDER
for _, row in deed_log.iterrows():
    if row["Status"] == "Complete":
        deed_file_name = str(row[deed_pk]) + ".pdf"
        move_to_status_folder(deed_folder, deed_file_name, row["Status"])

#RESULTS
print(f"SD Files copied: {sd_files_copied}")
print(f"Files Moved: {files_moved}")
print(f"Folders Created: {len(folders_created)}")
print(f"Work Request Files Moved: {wr_files_moved}")