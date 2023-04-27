import os

def purge_files(ip, files_to_purge, username, password):
    try:
        # Create the UNC path to the remote machine's C drive
        unc_path = f"\\\\{ip}\\C$"
        
        # Create a command to delete each file
        del_commands = [f"del /q \"{unc_path}\\{file_path}\"" for file_path in files_to_purge]
        
        # Combine the commands into a single command string and run it using os.system()
        full_command = " & ".join(del_commands)
        os.system(f"cmd /c {full_command}")
        
        return 0
    except Exception as e:
        print(f"Error purging files on {ip}: {e}")
        return 1
