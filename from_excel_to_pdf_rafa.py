import os
import pandas as pd
from fillpdf import fillpdfs

# Print the running file
print("RUNNING:", __file__)

# Paths and filenames
input_excel_path = "data/Database.xlsx"
input_folder = "data/PDF_samples"
output_directory = "Carpeta_PDF"

# Print status
print("READING DATA FROM", input_excel_path, "...")

# Read the data from the Excel file, starting from the fourth row (header=3)
df = pd.read_excel(input_excel_path, header=3)

# Lists of filenames to define specific conditions
ls_ja_off = ['E-2-Datenblatt-fuer-Erzeugungseinheiten copia.pdf',
             'E8-Inbetriebsetzungsprotokoll-fuer-Erzeugungsanlagen copia.pdf',
             'Inbetriebnahme_Inbetriebsetzung.pdf',
             'Inbetriebsetzung copia.pdf']
ls_1 = ['EVR-E3-Datenblatt-Speicher-V-2 copia.pdf']

# Default values for 'yes' and 'no' in the PDF forms
yes_default = "Yes"
no = "Off"

# Iterate over each file in the input folder
for file in os.listdir(input_folder):
    input_pdf_path = os.path.join(input_folder, file)

    # Adjust the 'yes' value based on the filename
    if file in ls_ja_off:
        yes = "Ja"
    elif file in ls_1:
        yes = "1"
    else:
        yes = yes_default

    print("\n\nREADING PDF DATA FROM", input_pdf_path)

    # Iterate over each row in the dataframe
    for index, row in df.iterrows():
        data_dic = row.to_dict()
        
        # Replace 'Y' with 'yes' and 'N' with 'no' in the data dictionary
        data_dic = {k: yes if v == "Y" else v for k, v in data_dic.items()}
        data_dic = {k: no if v == "N" else v for k, v in data_dic.items()}

        # Generate the output directory based on the 'Name' field
        name = str(data_dic['Name']).split()
        name = "_".join(name)
        output_directory_person = os.path.join(output_directory, name)

        # Create the output directory if it doesn't exist
        if not os.path.exists(output_directory_person):
            os.makedirs(output_directory_person)

        # Construct the output PDF path
        output_pdf_path = os.path.join(output_directory_person, f"output_{file}")

        # Print status
        print("WRITING IN", output_pdf_path)

        # Write the filled PDF
        fillpdfs.write_fillable_pdf(input_pdf_path=input_pdf_path, output_pdf_path=output_pdf_path, data_dict=data_dic)
