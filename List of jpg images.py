import os
from openpyxl import Workbook #to read excel files
from PIL import Image #to deal with images

def foto(file_path):
    try:
        # verify the file is an image by opening it as image
        Image.open(file_path).verify()
        return True
    except Exception as e:
        return False

def create_excel_list(folder_path, output_excel, extension='jpg'):
    # create a new excel file
    wb = Workbook()
    ws = wb.active

    for root, dirs, files in os.walk(folder_path):
        #files with *.jpg extension in the current folder
        jpg_files = [file for file in files if file.lower().endswith(f'.{extension}')]

        # for filtering out non-image files
        image_files = [file for file in jpg_files if foto(os.path.join(root, file))]

        # write data to excel
        for image_file in image_files:
            row_data = [image_file]
            ws.append(row_data)

    # save the excel file
    wb.save(output_excel)
    print(f"Excel list saved to {output_excel}")

# folder location:
x= input("Location of the input folder:")
y= str(input("Output file name:"))
input_folder = x
output_excel_file = f"{y}.xlsx"
create_excel_list(input_folder, output_excel_file, extension='jpg')
