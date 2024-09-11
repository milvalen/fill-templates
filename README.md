### fill-templates

#### Overview
This Python script automates the generation of `.docx` documents based on a template and data from multiple `.xlsx` files. The script replaces placeholders in the `.docx` template with values from the corresponding rows in the Excel files. The placeholders in the template are in the format `<<<index>>>`, where `index` refers to the column number (starting from 0) in the Excel file.

#### Key Features
- Replaces placeholders in `.docx` files based on column indices from Excel files.
- Processes all `.xlsx` files in a specified directory.
- Generates `.docx` documents for each row of data in the Excel file.
- The output file names are based on the first three columns of each data row.

#### Usage

1. Place the `.xlsx` files with data in a directory (default: `data`).
2. Create a `.docx` template file with placeholders like `<<<0>>>`, `<<<1>>>`, etc., where numbers refer to column indices.
3. Run the script. It will:
   - Find all Excel files in the `data` directory.
   - For each row in each Excel file, replace the placeholders in the template with values from the corresponding columns.
   - Save each generated document in the `results` directory. The file name will be based on the first three columns of the row.

#### Example

Suppose an Excel file has the following columns and data:
```
Name    Date       Event
John    2024-01-01 Marathon
Doe     2024-05-05 Triathlon
```

And the `.docx` template contains placeholders like `<<<0>>>`, `<<<1>>>`, `<<<2>>>`.

For each row, the script will generate `.docx` files:
- `John_2024-01-01_Marathon.docx`
- `Doe_2024-05-05_Triathlon.docx`

#### Execution

Ensure the necessary Python packages are installed:
```bash
python3 -m pip install openpyxl python-docx
```

Run the script from the command line or an IDE:
```bash
python3 main.py
```

#### Error Handling
- If no `.xlsx` files are found in the specified folder, a `FileNotFoundError` will be raised.
- If the template file is missing, the script will also raise a `FileNotFoundError`.
