import shutil
import os
import fitz
import base64
import requests
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Replace 'C:/Users/josh.brandt/OneDrive - Slalom/Saved/Jupyter' with the directory path you want to list the contents of
#source
directory_path = 'C:/Users/josh.brandt/OneDrive - Slalom/Saved/Jupyter/receipt_process/input_receipt'
#destincation
destination_folder = 'C:/Users/josh.brandt/OneDrive - Slalom/Saved/Jupyter/receipt_process/output_receipt'

# OpenAI API Key
api_key = "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"

# Define the scope and credentials
scope = ['https://www.googleapis.com/auth/spreadsheets']
creds = ServiceAccountCredentials.from_json_keyfile_name(r'C:\Users\josh.brandt\OneDrive - Slalom\Saved\Jupyter\credentials.json', scope)

# Open the spreadsheet by its ID
spreadsheet_id = '1qC-R4HslOBAomPlypJt1GVJVN5CMQeoa6OwlNOTrLUc'

#########################
#Initial Directory
#########################


def list_directory_contents(directory):
    file_list = []  # Initialize an empty list to store file names
    
    # Check if the directory exists
    if not os.path.exists(directory):
        print(f"The directory '{directory}' does not exist.")
        return file_list
    
    index = 0
    
    # List the files in the directory
    # print(f"Files in '{directory}':")
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        if os.path.isfile(file_path):
            
            list_dir_filepath = file_path
            list_dir_filename = filename.split('.')[0]
            list_dir_index = index
            list_dir_extension = os.path.splitext(file_path)[1]

            file_metadata = [f'{list_dir_filepath}',f'filename={list_dir_filename}',f'index={list_dir_index}',f'extension={list_dir_extension}']
            
            file_list.append(file_metadata)  # Append filename to the list
            index+=1
    
    return file_list  # Return the list of file names

files_array = list_directory_contents(directory_path)


#########################

print('input directory stored')

#########################
#Converting pdfs to jpgs
#########################


def pdf_to_jpg(pdf_path, output_path, dpi=300):
    # Open the PDF file
    pdf_document = fitz.open(pdf_path)

    # Iterate through each page
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        
        # Convert the page to an image with higher resolution (set DPI)
        image = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72))
        
        # Save the image as JPG
        image.save(f'{output_path}_page{page_num + 1}.jpg')

    # Close the PDF document
    pdf_document.close()

output_file_path = []  # Initialize an empty list to store paths for output files

for file_info in files_array:
    #Parses the current file path
    orig_file_path = file_info[0].split('=')[-1]
    #Parses the filename
    filename = file_info[1].split('=')[1]
    #joins the destination folder with the file name (sans extension)
    #extension is branch specific
    file_dest_path = os.path.join(destination_folder,filename)
    
    if file_info[3] == 'extension=.pdf':
    
        pdf_to_jpg(orig_file_path, file_dest_path, dpi=300)  # Adjust the DPI as needed
        output_file_path.append(file_dest_path+'_page1.jpg')
        
    elif file_info[3] == 'extension=.jpg':
        
        copy_file_extension = file_info[3].split('=')[1]
        copy_dest_path = file_dest_path+copy_file_extension
        
        shutil.copyfile(orig_file_path, copy_dest_path)
        output_file_path.append(copy_dest_path)
        
    elif file_info[3] == 'extension=.png':
        
        copy_file_extension = file_info[3].split('=')[1]
        copy_dest_path = file_dest_path+copy_file_extension
        
        shutil.copyfile(orig_file_path, copy_dest_path)    
        output_file_path.append(copy_dest_path)
    
    else:
        print('wrong format')

    #Removing the old files
    if os.path.exists(orig_file_path):
        os.remove(orig_file_path)
    else:
        print(f"File not found: {orig_file_path}")  


#########################

#print(output_file_path)
print('Copied and converted files to output folder')

#########################
#Chat GPT output
#########################

headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }

#client = OpenAI(api_key=api_key)

# Function to encode the image
def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

final_output_array = []

for output_file in output_file_path:
    #Parses the current file path
    #image_path = output_file[0].split('=')[-1]
    image_path = output_file
    
    # Getting the base64 string
    base64_image = encode_image(image_path)
    
    payload = {
        "model": "gpt-4-vision-preview",
        "messages": [
          {
            "role": "user",
            "content": [
              {
                "type": "text",
                  #"text": "Can you parse through the image of a receipt, and output the following as a single line. Please correct any potential typo of the vendor. Here is the desired format: \n [Date Created (in this format: YYY-MM-DD)]- [Vendor] - [Total Cost (in this format:Rounded to nearest Dollar, integer only] \n Example single line output: 2023-03-07 - Trader Joes - 200 \n Additionally, I would like to see the following: \n [Line Item 1] - [Item 1 cost] - [best guess of what the item is] - [best guess category of item. E.g. food, necessity, electronic, health, etc.]\n [Line Item 2] - [Item 2 cost] - [best guess of what the item is] - [best guess category of item. E.g. food, necessity, electronic, health, etc.] \n etc. \n Example: Rcla HyLmn - $9.98 - Ricola Honey Lemon - Medicine \n I would like the output to be as few words as possible. Only adhering to the format provided, and using the examples."
                  "text": "Can you parse through the image of a receipt, and output a python dictionary? I would like the array to have a header object called vendor_info that has the following attributes: \n [created (in this format: YYY-MM-DD)],[vendor_name],[total_cost] \n Example header - \"vendor_info\": { \"created\":\"2023-11-08\", \"vendor_name\":\"HOME DEPOT\", \"total_cost\":\"202.28\"} \n Additionally, I would like to see the individual receipt line items as a nested dictionary with these attributes: \n [Listed Line Item 1],[Item 1 Cost],[Best guess what the item is], [Best guess category of item. E.g. food, necessity, electronic, health, etc.] \n Example nested dictionary - \"line_item\": { \"receipt_item\":\"Rcla HyLmn\", \"cost\":\"9.98\", \"description\":\"Ricola Honey Lemon\", \"category\":\"Medicine\"} \n I would like the output to be as few words as possible. Do not deviate from the format provided. Make sure the cost items are down to the decimal. If you can't find line items, return  an array of objects with a single value called skipped:[reason]. For example, \"line_items\" : [{ \"skipped\":\"No visible line items\"}]. Please use double quotes everywhere."
                  #"text": "Can you parse through the image of a receipt, and output the following as a single line. Please correct any potential typo of the vendor. Here is the desired format: \n [Date Created (in this format: YYY-MM-DD)]- [Vendor] - [Total Cost (in this format:Rounded to nearest Dollar, integer only] \n Example single line output: 2023-03-07 - Trader Joes - 200 \n Additionally, I would like to see the following: \n [Line Item 1] |~| [Item 1 cost] |~| [best guess of what the item is] |~| [best guess category of item. E.g. food, necessity, electronic, health, etc.]\n [Line Item 2] |~| [Item 2 cost] |~| [best guess of what the item is] |~| [best guess category of item. E.g. food, necessity, electronic, health, etc.] \n etc. \n Example: Rcla HyLmn |~| $9.98 |~| Ricola Honey Lemon |~| Medicine \n I would like the output to be as few words as possible. Only adhering to the format provided, and using the examples."
              },
              {
                "type": "image_url",
                "image_url": {
                  "url": f"data:image/jpeg;base64,{base64_image}"
                }
              }
            ]
          }
        ],
        "max_tokens": 300
    }
    
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)

    # print("Initial JSON:")
    # print(response.json())
   
    # Extracting necessary data from the JSON response
    content = response.json()['choices'][0]['message']['content']
    
    # Extract JSON content within the triple-backtick code block
    start_idx = content.find("{")
    end_idx = content.rfind("}")
    if start_idx != -1 and end_idx != -1:
        extracted_content = content[start_idx:end_idx + 1]  # Extracting JSON content
    
        # print("Extracted Content:")
        # print(extracted_content)  # Troubleshooting print
        
        # Clean up the JSON-like content to make it valid JSON
        clean_content = (
            extracted_content.replace("\n", "")
            .replace(", }", "}")
            .replace(",]", "]")
        )
    
        # print("Cleaned Content:")
        # print(clean_content)  # Troubleshooting print
        
        try:
            # Load the cleaned content into a dictionary using json.loads
            receipt_data = json.loads(clean_content)
    
            line_items = receipt_data.get('line_items')
            # print(line_items)
            
            #if isinstance(line_items, str) and 'skipped' in line_items.lower():
            if any('skipped' in item for item in line_items):
                # Construct the line_items section based on the skipped scenario
                print('Line Items Skipped!')
                line_items = [
                    {
                        "skipped": "No visible line items"
                    }
                ]
                
                line_items = [
                    {
                        "receipt_item": receipt_data['vendor_info']['vendor_name'],
                        "cost": receipt_data['vendor_info']['total_cost'],
                        "description": f"{line_items}",
                        "category": 'Not Applicable'
                    }
                ]
            else:
                # Handle the regular line_items scenario
                line_items = receipt_data.get('line_items', [])
                print('Line Items Normal')
                
            # Modifying the data structure to fit the desired format
            receipt_info = {
                "vendor_info": {
                    "created": receipt_data['vendor_info']['created'],
                    "vendor_name": receipt_data['vendor_info']['vendor_name'],
                    "total_cost": receipt_data['vendor_info']['total_cost'],
                    "image_path": image_path
                },
                "line_items": line_items  # This will be the updated line_items based on the scenario
            }

            print(f"Analyzed: {receipt_data['vendor_info']['created']} - {receipt_data['vendor_info']['vendor_name']} - {receipt_data['vendor_info']['total_cost']}")
            final_output_array.append(receipt_info)
    
        except json.JSONDecodeError as e:
            print(f"JSONDecodeError: {e}")
            # Handle JSON decoding issues
    
    else:
        print("JSON content not found or invalid")


#########################

#print(final_output_array)
print('ChatGPT Analysis Complete')

######################### 
# Export to google sheet
#########################

client = gspread.authorize(creds)

# Open the spreadsheet by its ID
spreadsheet = client.open_by_key(spreadsheet_id)

# Select the specific worksheet within the spreadsheet using its index
worksheet_index = 0  # Change the index to select a different sheet if required
worksheet = spreadsheet.get_worksheet(worksheet_index)


for receipt in final_output_array:

    #Old version where I was excluding the image path from the header
    #receipt_header = {key: value for key, value in receipt['vendor_info'].items() if key != 'image_path'}
    
    #Receipt Header information
    receipt_header = receipt['vendor_info']
    #extracts the path to the image file
    old_receipt_path = receipt['vendor_info']['image_path']
    #extracts the line items into their own array
    receipt_items = receipt['line_items']

    #new receipt name format
    new_receipt_name = f"{receipt_header['created']} - {receipt_header['vendor_name']} - {receipt_header['total_cost']}"

    for line_item in receipt_items:
        
        # Separate variables
        item_name = line_item['receipt_item']
        item_price = float(line_item['cost'])
        item_description = line_item['description']
        item_category = line_item['category']
        
        # # Display separate variables
        # print(item_name)
        # print(item_price)
        # print(item_description)
        # print(item_category)
        
        # Add rows
        row_data = [new_receipt_name,item_name, item_price,item_description,item_category]
        # print(row_data)
        worksheet.append_row(row_data)

#######################
# Rename files 
#######################

    #new receipt name format
    filename_total_cost = float(receipt_header['total_cost'])
    filename_rounded_cost = round(filename_total_cost)
    
    new_receipt_name = f"{receipt_header['created']} - {receipt_header['vendor_name']} - {filename_rounded_cost}"
    
    # Get the directory and file extension from the old_receipt_path
    directory, file_extension = os.path.splitext(old_receipt_path)
    directory = os.path.dirname(directory)

    # Construct the new path with the updated file name
    new_receipt_path = os.path.join(directory, f"{new_receipt_name}{file_extension}")

    #Rename the file
    #Need to check to make sure it doesn't already exist
    if os.path.exists(new_receipt_path):
        
        # Counter to add differentiation
        counter = 1
        new_path = f"{new_receipt_name}_{counter}{file_extension}"
        
        # Check if the new path with counter exists, if so, increment counter
        while os.path.exists(new_path):
            counter += 1
            new_path = f"{new_receipt_name}_{counter}{file_extension}"
        
        # Now perform the renaming
        os.rename(old_receipt_path, new_path)
    else:
        os.rename(old_receipt_path, new_receipt_path)

print("Data added successfully!")