import os
import shutil
import socket
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
import zipfile
from datetime import datetime

REQUIRED_COLUMNS = {
    'Employee Number',
    'Employee Name',
    'Current Hire Date',
    'Work Country',
    'Business Title',
    'Email Address',
    'Business Group'
}

def find_matching_files(target_directory, valid_extensions):
    matching_files = []
    for root, _, files in os.walk(target_directory):
        for file in files:
            file_path = os.path.join(root, file)

            if file.endswith(valid_extensions):
                try:
                    matching_file = process_matching_file(file_path)
                    if matching_file:
                        matching_files.append(matching_file)
                except Exception as e:
                    print(f"Error reading file {file_path}: {e}. Skipping this file.")
                    continue

            elif file.endswith('.zip'):
                try:
                    with zipfile.ZipFile(file_path, 'r') as archive:
                        found_matching_file = False
                        for archive_file in archive.namelist():
                            with archive.open(archive_file) as af:
                                temp_path = os.path.join(target_directory, 'temp', archive_file)
                                parent = os.path.dirname(temp_path)
                                os.makedirs(parent, exist_ok=True)
                                with open(temp_path, "wb") as tf:
                                    shutil.copyfileobj(af, tf)
                                matching_file = process_matching_file(temp_path)
                                os.remove(temp_path)

                            if matching_file:
                                found_matching_file = True
                                break

                        if found_matching_file:
                            matching_files.append(file_path)

                except Exception as e:
                    print(f"Error processing archive {file_path}: {e}. Skipping this file.")
                    continue
                
    return matching_files

def process_matching_file(file_path):
    if file_path.endswith('.xls'):
        engine = 'xlrd'
        df = pd.read_excel(file_path, engine=engine, nrows=0)
    elif file_path.endswith('.xlsx'):
        engine = 'openpyxl'
        df = pd.read_excel(file_path, engine=engine, nrows=0)
    elif file_path.endswith('.csv'):
        df = pd.read_csv(file_path, nrows=0)
    else:
        return None

    if all(column in df.columns for column in REQUIRED_COLUMNS):
        return file_path
    return None

def copy_file_and_create_info_file(src, dest_folder, hostname, username, delete_after_copy=False):
    Path(dest_folder).mkdir(parents=True, exist_ok=True)

    dest_file = os.path.join(dest_folder, os.path.basename(src))
    shutil.copy(src, dest_file)

    if delete_after_copy:
        os.remove(src)

    info_filename = os.path.basename(src) + "_info.txt"
    info_filepath = os.path.join(dest_folder, info_filename)

    file_stat = os.stat(src)
    file_size_kb = file_stat.st_size / 1024
    file_creation_date = datetime.fromtimestamp(file_stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S')
    file_last_modified_date = datetime.fromtimestamp(file_stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
    file_last_accessed_date = datetime.fromtimestamp(file_stat.st_atime).strftime('%Y-%m-%d %H:%M:%S')
    
    with open(info_filepath, 'w') as f:
        f.write(f"File location: {src}\n")
        f.write(f"Hostname: {hostname}\n")
        f.write(f"Username: {username}\n")
        f.write(f"File size: {file_size_kb:.2f} KB\n")
        f.write(f"File creation date: {file_creation_date}\n")
        f.write(f"File last modified date: {file_last_modified_date}\n")
        f.write(f"File last accessed date: {file_last_accessed_date}\n")

def get_destination_folder(base_folder, hostname, username):
    return os.path.join(base_folder, f"{hostname}.{username}")

def main():
    user_folders = ['Documents', 'Downloads', 'Desktop', 'Box', 'OneDrive']
    shared_folder = '\\\\s-amusdat-ile03\\Cyber-Review\\GlobalR\\'
    Path(shared_folder).mkdir(parents=True, exist_ok=True)

    hostname = socket.gethostname()
    valid_extensions = ('.xls', '.xlsx', '.csv')
    users_path = 'C:\\Users'
    for user in os.listdir(users_path):
        user_path = os.path.join(users_path, user)
        if os.path.isdir(user_path):
            for user_folder in user_folders:
                target_directory = os.path.join(user_path, user_folder)
                if os.path.exists(target_directory):
                    for matching_file in find_matching_files(target_directory, valid_extensions):
                        dest_folder = get_destination_folder(shared_folder, hostname, user)
                        copy_file_and_create_info_file(matching_file, dest_folder, hostname, user)

if __name__ == "__main__":
    main()
