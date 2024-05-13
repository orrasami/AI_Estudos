from office365_api import SharePoint
import re
import sys, io
import pandas as pd
from pathlib import PurePath

# 1 args = SharePoint folder name. May include subfolders YouTube/2022
# FOLDER_NAME = sys.argv[1]
# 2 args = SharePoint file name. This is used when only one file is being downloaded
# If all files will be downloaded, then set this value as "None"
# FILE_NAME = sys.argv[2]
# 3 args = SharePoint file name pattern
# If no pattern match files are required to be downloaded, then set this value as "None"
# FILE_NAME_PATTERN = sys.argv[3]


def modify_file(file_obj):
    data = io.BytesIO(file_obj)
    ws_list = pd.ExcelFile(data).sheet_names
    df = pd.read_excel(data, sheet_name=ws_list[0])

    # apply changes
    df['new column'] = "Testing Column"

    # Create excel object in memory
    output_obj = io.BytesIO()
    writer = pd.ExcelWriter(output_obj)
    df.to_excel(writer, index=0, sheet_name=ws_list[0])
    writer.save()
    output_obj.seek(0)

    upload_file('Testing Modified File', FOLDER_NAME, output_obj.read())


def upload_file(file_name, folder_name, content):
    SharePoint().upload_file(file_name, folder_name, content)


def get_file(file_n, folder):
    file_obj = SharePoint().download_file(file_n, folder)
    modify_file(file_obj)


def get_files(folder):
    files_list = SharePoint()._get_files_list(folder)
    for file in files_list:
        get_file(file.name, folder)


def get_files_by_pattern(keyword, folder):
    files_list = SharePoint()._get_files_list(folder)
    for file in files_list:
        if re.search(keyword, file.name):
            get_file(file.name, folder)


if __name__ == '__main__':
    FILE_NAME = "Python.xlsx"
    FOLDER_NAME = "General/5. ATAS"
    FILE_NAME_PATTERN = None

    if FILE_NAME != 'None':
        get_file(FILE_NAME, FOLDER_NAME)
    elif FILE_NAME_PATTERN != 'None':
        get_files_by_pattern(FILE_NAME_PATTERN, FOLDER_NAME)
    else:
        get_files(FOLDER_NAME)