import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import json
from openpyxl import load_workbook, Workbook
import os
import time

# Script for uploading files to Pure and updating research outputs
# Author: Maria Camila Hernández

# Read the list of file URLs, UUIDs, and filenames from the Excel file
input_excel_file = 'Haplo_files_pure/files_to_upload_pure.xlsx'  # Change this to your input file name
df = pd.read_excel(input_excel_file)

# Your Basic Auth credentials for downloading files
download_username = 'API_Username'  # Replace with your download API username
download_password = 'API_Password'  # Replace with your download API password

# Your API key for uploading files to Pure system
pure_api_key = 'Pure_API_Key'  # Replace with your Pure API key

# Output Excel file for response data
output_excel_file = 'Haplo_files_pure/upload_response.xlsx'  # Change this to your output file name

# Function to append a row to the Excel file
def append_to_excel(file_name, data):
    try:
        if os.path.exists(file_name):
            # Load an existing workbook
            book = load_workbook(file_name)
            sheet = book.active
        else:
            # Create a new workbook if the file does not exist
            book = Workbook()
            sheet = book.active
            sheet.append(list(data.keys()))  # Append the header

        sheet.append(list(data.values()))  # Append the data
        book.save(file_name)
    except Exception as e:
        print(f"An error occurred while trying to append to the Excel file: {str(e)}")

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    file_url = row['url']  # Assuming the file URL is in a column named 'url'
    uuid = row['UUID']  # Assuming the UUID is in a column named 'UUID'
    file_name = row['filename']  # Assuming the filename is in a column named 'filename'
    
    try:
        # Step 1: Download the file from the provided URL with Basic Authentication
        file_response = requests.get(file_url, auth=HTTPBasicAuth(download_username, download_password), verify=False, timeout=60)
        file_response.raise_for_status()
        file_content = file_response.content
        print(f"Successfully downloaded file from {file_url}")
        
        # Step 2: Upload the file to the Pure system API
        upload_url = 'https://Your_URL/ws/api/research-outputs/file-uploads'
        headers = {
            'accept': 'application/json',
            'api-key': pure_api_key
        }
        upload_response = requests.put(upload_url, headers=headers, data=file_content, timeout=60)
        upload_response.raise_for_status()
        upload_data = upload_response.json()
        print(f"Successfully uploaded file from {file_url} to Pure system")

         # Add a delay between requests to Elsevier
        time.sleep(2)  # Delay for 5 seconds
        
        # Step 3: Make the second request to update the research output
        update_url = f'https://Your_URL/ws/api/research-outputs/{uuid}'
        update_headers = {
            'accept': 'application/json',
            'api-key': pure_api_key,
            'content-type': 'application/json'
        }
        update_payload = {
            "electronicVersions": [
                {
                    "typeDiscriminator": "FileElectronicVersion",
                    "accessType": {
                        "uri": "/dk/atira/pure/core/openaccesspermission/open",
                        "term": {
                            "en_GB": "Open"
                        }
                    },
                    "file": {
                        "fileName": file_name,
                        "size": upload_data['size'],
                        "mimeType": upload_data['mimeType'], 
                        "uploadedFile": {
                            "digest": upload_data['digest'],
                            "digestType": upload_data['digestType'],
                            "size": upload_data['size'],
                            "mimeType": upload_data['mimeType'],  
                            "timeStamp": upload_data['timeStamp'],
                            "expires": upload_data['expires'],
                            "key": upload_data['key']
                        }
                    }
                }
            ]
        }
        update_response = requests.put(update_url, headers=update_headers, data=json.dumps(update_payload), timeout=60)
        update_response.raise_for_status()
        print(f"Successfully updated research output for UUID {uuid}")
        
        # Prepare the response data
        response_data = {
            'url': file_url,
            'uuid': uuid,
            'digest': upload_data['digest'],
            'digestType': upload_data['digestType'],
            'size': upload_data['size'],
            'mimeType': upload_data['mimeType'],
            'timeStamp': upload_data['timeStamp'],
            'expires': upload_data['expires'],
            'key': upload_data['key']
        }

        # Append the response data to the Excel file
        append_to_excel(output_excel_file, response_data)
        print(f"Response data saved to {output_excel_file}")
    except requests.exceptions.RequestException as e:
        print(f"Failed to process file from {file_url}: {str(e)}")
        error_data = {
            'url': file_url,
            'uuid': uuid,
            'error': str(e)
        }
        append_to_excel(output_excel_file, error_data)
        print(f"Error data saved to {output_excel_file}")
        