# PDF-filler-
Program to fill PDF files from Excel data

## How to use

- First time (or with dependencies changed):
```bash
bash dependencies.sh
```
- To run the program:
```bash
bash run.sh
```

## Paths

Inside the `run.sh` file, you can change the paths to the input files and the output folder.
Also, you can copy and paste the calls to `main.py` to run the program with different files.

## Command line running

```bash
source .venv/bin/activate
```

```bash
python main.py --pdf-file "path/to/pdf/file" --excel-file "path/to/excel/file" --output-folder "path/to/output/folder"
```

Note: the paths can be absolute or relative, but beware of where the program is being run from.