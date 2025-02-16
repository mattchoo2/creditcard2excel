# How to Convert PDF to Excel Using PowerShell or Command Prompt

## Prerequisites
Before proceeding, ensure you have the following installed on your system:
- Python (latest version)
- `pdfplumber` Python library
- PowerShell or Command Prompt

## Step 1: Install Python
1. Open PowerShell or Command Prompt.
2. Download and install Python:
   - Visit [Python's official website](https://www.python.org/downloads/) and download the latest version.
   - Run the installer and ensure you select **"Add Python to PATH"** before clicking **Install Now**.
3. Verify installation by running:
   ```sh
   python --version
   ```

## Step 2: Install `pdfplumber`
1. Open PowerShell or Command Prompt.
2. Install `pdfplumber` using pip:
   ```sh
   pip install pdfplumber pandas openpyxl
   ```

## Step 3: Prepare PDF Files
1. Place the PDF files you want to convert in the `Downloads` folder.
   - Example path: `C:\Users\YourUsername\Downloads`

## Step 4: Run the Script to Convert PDF to Excel
1. Open Notepad and paste the following Python script:
   ```python
   import pdfplumber
   import pandas as pd
   import os

   input_folder = os.path.expanduser("~/Downloads")
   output_folder = os.path.expanduser("~/Downloads")

   for file in os.listdir(input_folder):
       if file.endswith(".pdf"):
           pdf_path = os.path.join(input_folder, file)
           output_excel = os.path.join(output_folder, file.replace(".pdf", ".xlsx"))
           
           data = []
           with pdfplumber.open(pdf_path) as pdf:
               for page in pdf.pages:
                   table = page.extract_table()
                   if table:
                       data.extend(table)
           
           df = pd.DataFrame(data)
           df.to_excel(output_excel, index=False)
           print(f"Converted {file} to Excel successfully!")
   ```
2. Save the file as `convert_pdf_to_excel.py` in the `Downloads` folder.
3. Open PowerShell or Command Prompt and navigate to `Downloads`:
   ```sh
   cd %USERPROFILE%\Downloads
   ```
4. Run the script:
   ```sh
   python convert_pdf_to_excel.py
   ```

## Step 5: Check the Output
- The converted Excel files will be saved in the `Downloads` folder.
- Open the `.xlsx` file using Excel to verify the extracted data.

## Troubleshooting
- If Python is not recognized, restart your computer or reinstall Python ensuring **"Add Python to PATH"** is selected.
- If `pdfplumber` is missing, try reinstalling it with:
  ```sh
  pip install --upgrade pdfplumber pandas openpyxl
  ```
- If the script doesn't work for your PDFs, some documents may have embedded images instead of selectable text.

By following these steps, you can efficiently extract tables from PDFs into Excel using PowerShell or Command Prompt with Python.

