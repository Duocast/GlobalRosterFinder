import os
import shutil
import socket
from pathlib import Path
import csv

REQUIRED_COLUMNS = {
    'employee number',
    'employee name',
    'current hire date',
    'work country',
    'business title',
    'email address',
    'business group'
}

def find_matching_files(start_directory, valid_extensions=None):
    if valid_extensions is None:
        valid_extensions = ['.csv']

    def contains_required_columns(columns):
        columns_set = {col.strip().lower().replace('\ufeff', '') for col in columns}
        return REQUIRED_COLUMNS.issubset(columns_set)

    for root, _, files in os.walk(start_directory):
        for file in files:
            file_extension = os.path.splitext(file)[1]
            if file_extension.lower() in valid_extensions:
                file_path = os.path.join(root, file)
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as csvfile:
                        print(f"Processing file: {file_path}")
                        reader = csv.reader(csvfile)
                        header = next(reader)
                        if contains_required_columns(header):
                            print(f"Found matching file: {file_path}")
                            yield file_path
                        else:
                            missing_columns = REQUIRED_COLUMNS - {col.strip().lower() for col in header}
                            print(f"File {file_path} does not contain the required columns: {missing_columns}")
                            print(f"Header in file: {header}")
                except Exception as e:
                    print(f"Error reading file {file_path}: {e}")

def copy_file_and_create_info_file(src, dest_folder, hostname, username):
    Path(dest_folder).mkdir(parents=True, exist_ok=True)

    dest_file = os.path.join(dest_folder, os.path.basename(src))
    shutil.copy(src, dest_file)

    info_filename = os.path.basename(src) + "_info.txt"
    info_filepath = os.path.join(dest_folder, info_filename)
    with open(info_filepath, 'w') as f:
        f.write(f"File location: {src}\n")
        f.write(f"Hostname: {hostname}\n")
        f.write(f"Username: {username}\n")

def get_destination_folder(base_folder, hostname, username):
    return os.path.join(base_folder, f"{hostname}.{username}")

def main():
    user_folders = ['Documents', 'Downloads', 'Desktop']
    shared_folder = '\\\\s-amusdat-ile03\\Cyber-Review\\'
    Path(shared_folder).mkdir(parents=True, exist_ok=True)

    hostname = socket.gethostname()
    
    valid_extensions = ['.csv']
    users_path = 'C:\\Users'
    for user in os.listdir(users_path):
        user_path = os.path.join(users_path, user)
        if os.path.isdir(user_path):
            for user_folder in user_folders:
                target_directory = os.path.join(user_path, user_folder)
                if os.path.exists(target_directory):
                    for matched_file in find_matching_files(target_directory, valid_extensions):
                        print(f"Processing matched_file: {matched_file}")  # Added
                        try:
                            dest_folder = get_destination_folder(shared_folder, hostname, user)
                            copy_file_and_create_info_file(matched_file, dest_folder, hostname, user)
                        except Exception as e:
                            print(f"Error copying file {matched_file} to {dest_folder}: {e}")

if __name__ == "__main__":
    main()
