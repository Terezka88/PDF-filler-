
print("RUNNING:", __file__)
from fillpdf import fillpdfs
import pandas as pd 
import os 
import argparse

parser = argparse.ArgumentParser(description='Fill PDFs with data from Excel')
parser.add_argument('--input-excel-file', type=str, help='Name of the input Excel file')
parser.add_argument('--input-pdf-file', type=str, help='Name of the input PDF file')
parser.add_argument('--output-dir', type=str, default="output", help='Name of the output directory')

input_excel_file = parser.parse_args().input_excel_file
input_pdf_file = parser.parse_args().input_pdf_file
output_directory = parser.parse_args().output_dir
output_directory = os.path.join("outputs", output_directory)

print("Reading Excel data from:", input_excel_file, "\n\n")


if output_directory not in os.listdir():
    os.makedirs(output_directory)


# form_fields = list(fillpdfs.get_form_fields("data/sample.pdf").keys()) # to get the names of the boxes that we need to fill 
df = pd.read_excel(input_excel_file, header = 3) # we import the data 

print("Reading pdf data from:", input_pdf_file, "\n\n")
for index, row in df.iterrows():

    data_dic = row.to_dict()
    data_dic = {k: "Yes" if v == "Y" else v for k, v in data_dic.items()}
    data_dic = {k: "Off" if v == "N" else v for k, v in data_dic.items()}

    # print(data_dic)
    filename_ext = "_".join(data_dic['Text1'].split())
    output_pdf_path = os.path.join(output_directory, f"output_{filename_ext}.pdf" )  # Construct the output PDF path dynamically

    # Call the write_fillable_pdf function with the correct arguments
    print("WRriting in:", output_pdf_path)
    fillpdfs.write_fillable_pdf(input_pdf_path=input_pdf_file, output_pdf_path=output_pdf_path, data_dict=data_dic)