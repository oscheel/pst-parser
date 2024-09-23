import pypff
import os
import dataiku
from email import message_from_string

# Function to save email content to a .txt file
def save_email_as_txt(subject, body, output_dir, email_count):
    filename = f"email_{email_count}.txt"
    filepath = os.path.join(output_dir, filename)

    # Write email content to a text file
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(f"Subject: {subject}\n\n{body}")

# Function to process individual messages
def process_message(message, output_dir, email_count):
    try:
        subject = message.subject if message.subject else "No Subject"
        body = message.plain_text_body if message.plain_text_body else "No Content"
        
        save_email_as_txt(subject, body, output_dir, email_count)
    except Exception as e:
        print(f"Failed to process message: {str(e)}")

# Function to recursively process folders
def process_folder(folder, output_dir):
    email_count = 0
    for item in folder.sub_messages:
        process_message(item, output_dir, email_count)
        email_count += 1
    
    # Process subfolders recursively
    for subfolder in folder.sub_folders:
        process_folder(subfolder, output_dir)

# Main function to convert PST to text files
def convert_pst_to_txt(pst_file_path, output_dir):
    pst = pypff.file()
    pst.open(pst_file_path)
    
    root_folder = pst.get_root_folder()
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    process_folder(root_folder, output_dir)

    print(f"Conversion completed. Emails saved to {output_dir}")

# Function to download PST file from Dataiku folder and process it
def process_pst_from_dataiku(dataiku_folder, pst_filename, output_dir):
    # Connect to the Dataiku folder
    folder = dataiku.Folder(dataiku_folder)
    
    # Get the full path of the PST file from the managed folder
    pst_file_path = os.path.join('/mnt/dataiku_managed_folder/', pst_filename)
    
    # Download the PST file to a local path
    with folder.get_download_stream(pst_filename) as f:
        with open(pst_file_path, 'wb') as local_file:
            local_file.write(f.read())
    
    # Convert the downloaded PST file to TXT files
    convert_pst_to_txt(pst_file_path, output_dir)

# Example usage in Dataiku:
# Define the managed folder ID and PST filename in Dataiku
dataiku_folder = 'your_dataiku_managed_folder_id'
pst_filename = 'your_pst_file.pst'
output_dir = '/path_to_output_directory'

# Process the PST file
process_pst_from_dataiku(dataiku_folder, pst_filename, output_dir)
