source .venv/bin/activate

python main.py --input-excel-file data/Rafa_pdf.xlsx \
               --input-pdf-file data/template.pdf \
               --output-dir warstrugen 

python main.py --input-excel-file data/Rafa_pdf.xlsx \
               --input-pdf-file data/template.pdf \
               --output-dir pifosten

python main.py --input-excel-file data/Rafa_pdf.xlsx \
               --input-pdf-file data/template.pdf \
               --output-dir orta