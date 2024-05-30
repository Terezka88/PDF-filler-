import os
import pandas as pd
from fillpdf import fillpdfs

print("RUNNING:", __file__)

input_excel_path = "data/Rafa_pdf 2.xlsx"
input_folder = "data/PDF_samples"
output_directory = "Carpeta_PDF"

# Reading data from Excel file
print("READING DATA FROM", input_excel_path, "...")
try:
    df = pd.read_excel(input_excel_path, header=3)
except FileNotFoundError:
    print(f"Error: The file {input_excel_path} was not found.")
    exit(1)

# Create output directory if it doesn't exist
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Iterating over PDF files in the input folder
for file in os.listdir(input_folder):
    input_pdf_path = os.path.join(input_folder, file)
    #if input_pdf_path == 'data/PDF_samples/EVR-E3-Datenblatt-Speicher-V-2.pdf copia':
        #print(f"Skipping file {input_pdf_path}")
        #continue

    print("READING PDF DATA FROM", input_pdf_path)

    # Iterating over rows in the DataFrame
    for index, row in df.iterrows():
        data_dic = row.to_dict()
        # Applying data transformations
        data_dic = {k: ("Yes" if v == "Y" else v) for k, v in data_dic.items()}
        data_dic = {k: ("Off" if v == "N" else v) for k, v in data_dic.items()}

        # Constructing output file path
        filename_ext = "_".join(str(data_dic['Name']).split())
        output_pdf_path = os.path.join(output_directory, f"output_{filename_ext}_{file[:-4]}.pdf")

        # Writing data to PDF
        print("WRITING TO", output_pdf_path)
        try:
            fillpdfs.write_fillable_pdf(input_pdf_path=input_pdf_path, output_pdf_path=output_pdf_path, data_dict=data_dic)
        except KeyError as e:
            print(f"KeyError: {e} - Retrying with different data transformations.")
            data_dic = {k: ("Ja" if v == "Y" else v) for k, v in row.to_dict().items()}
            data_dic = {k: ("0" if v == "N" else v) for k, v in data_dic.items()}
            fillpdfs.write_fillable_pdf(input_pdf_path=input_pdf_path, output_pdf_path=output_pdf_path, data_dict=data_dic)
        except Exception as e:
            print(f"An error occurred: {e}")

