import os
import shutil
def create_folder(folder_name):
    try:
        # Create the folder if it doesn't exist
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
            return True
        else:
            return False
    except Exception as e:
        print(f"An error occurred: {e}")

def delete_folder(folder_name):
    try:
        if os.path.exists(folder_name):
            shutil.rmtree(folder_name)
    except Exception as e:
        print(f"An error occurred: {e}")

def create_textfile(folder_name, filename, data):
    try:
        # Create the folder if it doesn't exist
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
        # Create the file path
        file_path = os.path.join(folder_name, filename)
        # Write data to the file
        with open(file_path, 'w') as file:
            file.write(data)
        print(f"File '{filename}' created successfully in folder '{folder_name}'.")
    except Exception as e:
        print(f"An error occurred: {e}")

def delete_file(filename):
    try:
        os.remove(filename)
        print(f"File '{filename}' deleted.")
    except Exception as e:
        print(f"An error occurred: {e}")