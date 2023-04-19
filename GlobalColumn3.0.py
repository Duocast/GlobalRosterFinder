import os
import shutil
import socket
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

REQUIRED_COLUMNS = {
    'Employee Number',
    'Employee Name',
    'Current Hire Date',
    'Work Country',
    'Business Title',
    'Email Address',
    'Business Group'
}

def find_matching_files(start_directory, valid_extensions=None):
    if valid_extensions is None:
        valid_extensions = ['.csv', '.xls', '.xlsx']

    def contains_required_columns(columns):
        return REQUIRED_COLUMNS.issubset(set(columns))

    for root, _, files in os.walk(start_directory):
        for file in files:
            file_extension = os.path.splitext(file)[1]
            if file_extension.lower() in valid_extensions:
                file_path = os.path.join(root, file)
                try:
                    if file_extension.lower() == '.csv':
                        df = pd.read_csv(file_path, encoding='utf-8', errors='ignore', nrows=0)
                    else:
                        df = pd.read_excel(file_path, engine='openpyxl', nrows=0)

                    if contains_required_columns(df.columns):
                        yield file_path
                except Exception as e:
                    print(f"Error reading file {file_path}: {e}")

def copy_file_and_create_info_file(src, dest_folder, hostname, username):
    Path(dest_folder).mkdir(parents=True, exist_ok=True)

    dest_file = os.path.join(dest_folder, os.path.basename(src))
    shutil.copy(src, dest_file)

    info_filename = os.path.basename(src) + "_info.txt"
    info_filepath = os.path.join(dest_folder, info_filename)
    
    file_stat = os.stat(src)
    file_size = file_stat.st_size
    file_creation_date = file_stat.st_ctime
    file_last_modified_date = file_stat.st_mtime
    file_last_accessed_date = file_stat.st_atime

    with open(info_filepath, 'w') as f:
        f.write(f"File location: {src}\n")
        f.write(f"Hostname: {hostname}\n")
        f.write(f"Username: {username}\n")
        f.write(f"File size: {file_size} bytes\n")
        f.write(f"File creation date: {os.path.getctime(src)}\n")
        f.write(f"File last modified date: {os.path.getmtime(src)}\n")
        f.write(f"File last accessed date: {os.path.getatime(src)}\n")

def get_destination_folder(base_folder, hostname, username):
    return os.path.join(base_folder, f"{hostname}.{username}")

def main():
    user_folders = ['Documents', 'Downloads', 'Desktop', 'Box', 'OneDrive']
    shared_folder = '\\\\s-amusdat-ile03\\Cyber-Review\\'
    Path(shared_folder).mkdir(parents=True, exist_ok=True)

    hostname = socket.gethostname()
    valid_extensions = ['.csv', '.xls', '.xlsx']
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
