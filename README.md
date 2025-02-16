# How to Batch Convert Credit Card Statement PDF to Excel 
Note: due to `pdfplumber` imperfection in detecting blank cells. You do need to cut and align the date and value columns with the text labels.

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
   - Example path: `C:\Users\YourUsername\Downloads\A`

## Step 4: Run the Script to Convert PDF to Excel
1. Open Notepad and paste the following Python script:
   ***(Remember update folder path & password if any)***
   ```import pdfplumber
      import pandas as pd
      import os
      
      # Folder containing PDFs & file password if any
      pdf_folder = r"C:\Users\YourUsername\Downloads\A"
      password = "xxxx" 
      
      # Get all PDF files in the folder
      pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]
      
      for pdf_file in pdf_files:
          pdf_path = os.path.join(pdf_folder, pdf_file)
          output_file = os.path.join(pdf_folder, pdf_file.replace(".pdf", ".xlsx"))
          
          data = []
          max_columns = 0  # Track max columns for consistency
      
          try:
              with pdfplumber.open(pdf_path, password=password) as pdf:
                  print(f"✅ Opened PDF: {pdf_path} ({len(pdf.pages)} pages detected)")
      
                  for page_number, page in enumerate(pdf.pages, start=1):
                      print(f"🔹 Processing Page {page_number}")
      
                      tables = page.extract_tables()
                      if not tables:
                          print(f"⚠️ No structured table found on Page {page_number}, trying text extraction...")
                          page_text = page.extract_text()
                          if page_text:
                              tables = [page_text.split("\n")]  # Treat raw text lines as table rows
      
                      if tables:
                          print(f"✅ {len(tables)} table(s) found on Page {page_number}")
                          for table in tables:
                              for row in table:
                                  if row is None:
                                      empty_row = [""] * max_columns if max_columns > 0 else [""]
                                      data.append(empty_row)
                                      continue
      
                                  while len(row) < max_columns:
                                      row.append("")
      
                                  split_rows = []
                                  row_max_lines = 1
      
                                  for cell in row:
                                      if cell:
                                          lines = cell.strip().split("\n")
                                          split_rows.append(lines)
                                          row_max_lines = max(row_max_lines, len(lines))
                                      else:
                                          split_rows.append([""])
      
                                  for i in range(row_max_lines):
                                      new_row = [col[i] if i < len(col) else "" for col in split_rows]
                                      data.append(new_row)
      
                                  max_columns = max(max_columns, len(row))
      
          except Exception as e:
              print(f"❌ Failed to process {pdf_file}: {e}")
              continue
      
          for i in range(len(data)):
              while len(data[i]) < max_columns:
                  data[i].append("")
      
          df = pd.DataFrame(data)
          df.columns = [f"Column {i+1}" for i in range(max_columns)]
          df.to_excel(output_file, index=False)
          print(f"✅ Conversion done! Excel file saved at: {output_file}")
   ```
2. Save the file as `convert_pdf_to_excel.py` in the `Downloads\A` folder.
3. Open PowerShell or Command Prompt and navigate to `Downloads\A`:
   ```sh
   cd %USERPROFILE%\Downloads\A
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

