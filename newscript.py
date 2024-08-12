# -*- coding: utf-8 -*-
"""
Created on Tue Jun 25 14:36:38 2024

@author: kburg
"""
import sys
import pandas as pd
import os
import zipfile
import keras_ocr
import traceback

def process_image_from_excel(row_number, output_folder, input_file):
    # Read the Excel file
    df = pd.read_excel(input_file)

    # Assuming the row_number is provided as an index, adjust accordingly
    if row_number >= len(df):
        raise ValueError(f"Row number {row_number} is out of range.")

    # Get information from the specified row
    leaflet_path = df.loc[row_number, 'leaflet_path']
    image_file = df.loc[row_number, 'image_file']  # Assuming 'image_file' is the column name for image file
    folder_name = df.loc[row_number, 'folder_name']

    # Process image with keras_ocr pipeline
    leaflet_number = os.path.basename(leaflet_path).split('-')[-1].split('.')[0]
    pipeline = keras_ocr.pipeline.Pipeline()

    with zipfile.ZipFile(leaflet_path, 'r') as zip_ref:
        with zip_ref.open(image_file) as img_file:
            keras_image = keras_ocr.tools.read(img_file)
            predictions = pipeline.recognize([keras_image])[0]

    text = ' '.join([text for text, _ in predictions])

    # Save the output with the specified row number
    output_filename = os.path.join(output_folder, f"output_{row_number}.txt")
    with open(output_filename, 'w') as f:
        f.write(f"Leaflet Number: {leaflet_number}, Image: {image_file}, Text: {text}, Folder Name: {folder_name}")

    print(f"Processed image and saved output to '{output_filename}'")
    
    

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script_name.py <sysvariable>")
        sys.exit(1)
    
    sysvariable = int(sys.argv[1])  # Convert command-line argument to integer
    input_file = 'C:/Users/kburg/OneDrive/Documents/Trinity/diss_all_other/xcel.xlsx'
    output_folder = 'C:/Users/kburg/OneDrive/Documents/Trinity/diss_all_other'  # Replace with your desired output folder

    start_index = (sysvariable) * 1
    end_index = sysvariable *1

    df = pd.read_excel(input_file)

   # for i in range(start_index, end_index):
  #      try:
   #         process_image_from_excel(i, output_folder, input_file)
    #    except Exception as e:
     #       print(f"Error processing row {i}: {str(e)}")
      #      traceback.print_exc()  # Print the traceback
       #     continue
    for i in range(start_index, end_index):
        process_image_from_excel(i, output_folder, input_file)

        
#sysvariable= 3187  #user input will be the task array number from 1-200 or wtv your nubmer of jobs divded by 50 plus one is
 
#for i in range((sysvariable-1)*1,(sysvariable)*1): #this counts from 0 to 49, them 50 to 99, etc depending on what sysvariable is, test it it works
 #   print(i)
    
#18125/4


#9559/3


#18125/50
#363

# test for end of thing 6042


input_file = 'C:/Users/kburg/OneDrive/Documents/Trinity/diss_all_other/xcel.xlsx'
output_folder = 'C:/Users/kburg/OneDrive/Documents/Trinity/diss_all_other'  # Replace with your desired output folder


process_image_from_excel(4331, output_folder, input_file)

#2880 minutes per job

#40 min for each 50

50/40
1.25*75

2880/40

2880/4
720/50

9560/4

#needs 80 minutes to run 100

2880/20
144-1
144+143
287+143
#1-72 