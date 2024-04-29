
print("RUNNING:", __file__)
from fillpdf import fillpdfs
import pandas as pd 
import os 

input_excel_path = "data/Rafa_pdf.xlsx"
print("READING DATA FROM", input_excel_path, "...")
# form_fields = list(fillpdfs.get_form_fields("data/sample.pdf").keys()) # to get the names of the boxes that we need to fill 
df = pd.read_excel(input_excel_path, header = 3) # we import the data 

input_pdf_path = "data/sample.pdf"
output_directory = "Carpeta_PDF"

print("READING PDF DATA FROM", input_pdf_path)

if output_directory not in os.listdir():
    os.makedirs(output_directory)

for index, row in df.iterrows():

    data_dic = row.to_dict()
    data_dic = {k: "Yes" if v == "Y" else v for k, v in data_dic.items()}
    data_dic = {k: "Off" if v == "N" else v for k, v in data_dic.items()}

    # print(data_dic)
    filename_ext = "_".join(data_dic['Text1'].split())
    output_pdf_path = os.path.join(output_directory, f"output_{filename_ext}.pdf" )  # Construct the output PDF path dynamically

    # Call the write_fillable_pdf function with the correct arguments
    print("WRITING IN ", output_pdf_path)
    fillpdfs.write_fillable_pdf(input_pdf_path=input_pdf_path, output_pdf_path=output_pdf_path, data_dict=data_dic)