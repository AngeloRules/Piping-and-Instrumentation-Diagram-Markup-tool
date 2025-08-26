from pdf2image import convert_from_path
import constants
import os

dir_path = constants.pdfs_location
AllFiles = [file for file in os.listdir(dir_path) if file.endswith(".pdf")]
for file_name in AllFiles:
    file_path = os.path.join(dir_path,file_name)
    pages = convert_from_path(file_path)
    for i,page in enumerate(pages):
        page.save(f'{str(file_name)}_{i}.jpg', 'JPEG')