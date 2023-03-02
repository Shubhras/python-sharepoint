from office365_api import SharePoint
import re
import sys, os
from pathlib import PurePath
import datetime

print('PLEASE ENTER LOCAL FILE PATH:')
local_file_path = input()
print('DO YOU WANT TO CREATE A NEW FOLDER YOURSELF ? Y/N:')
folder_permission = input()
if folder_permission == "Y" or folder_permission == "y":
    print('PLEASE ENTER SHAREPOINT FOLDER NAME:')
    SHAREPOINT_FOLDER_NAME = input()



def upload_files(folder, keyword=None):
    if folder_permission == "Y" or folder_permission == "y":
        SharePoint().create_self_folder_upload_folder_to_sharepoint(folder, SHAREPOINT_FOLDER_NAME,local_file_path)
    else:
        SharePoint().upload_folder_to_sharepoint(folder,local_file_path)
def get_list_of_files(folder):
    file_list = []
    folder_item_list = os.listdir(folder)
    for item in folder_item_list:
        item_full_path = PurePath(folder, item)
        if os.path.isfile(item_full_path):
            file_list.append([item, item_full_path])
    return file_list

# read files and return the content of files
def get_file_content(file_path):
    with open(file_path, 'rb') as f:
        return f.read()

if __name__ == '__main__':
    # upload_files(ROOT_DIR, FILE_NAME_PATTERN)
    upload_files(local_file_path)
