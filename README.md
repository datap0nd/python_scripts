# Excel File Cloner

Reads data from an open Excel workbook via Windows COM (from memory) and writes it into a pre-existing xlsx template — without ever saving through Excel itself.

## How it works

1. Connects to the running Excel instance via `pywin32` (COM automation)
2. Duplicates an unzipped `.xlsx` template folder
3. Reads all cell values from the open workbook (bulk read from memory)
4. Builds the worksheet XML files and writes them into the duplicated template
5. Zips the folder into a new `.xlsx` file

Excel never performs a "Save" operation. The output file is assembled entirely by Python using standard `zipfile` and `xml.etree` — no third-party libraries touch the file.

## Requirements

- Windows
- Python 3.8+
- An open Excel instance with at least one workbook

## Installation

```bash
git clone https://github.com/datap0nd/python_scripts.git
cd python_scripts
pip install pywin32
```

## Template setup (one-time)

The script needs an unzipped `.xlsx` folder to use as a template. This provides the base XML structure (styles, themes, content types).

1. Take any `.xlsx` file (can be blank or a sample with your preferred formatting)
2. Rename it from `.xlsx` to `.zip`
3. Extract it to: `C:\Users\r.cunha\AppData\Local\Temp\xlsx_template`

You should end up with this structure:

```
xlsx_template/
├── [Content_Types].xml
├── _rels/
│   └── .rels
├── xl/
│   ├── workbook.xml
│   ├── styles.xml
│   ├── sharedStrings.xml
│   ├── worksheets/
│   │   └── sheet1.xml
│   └── theme/
│       └── theme1.xml
└── docProps/
    ├── app.xml
    └── core.xml
```

To change the template location, edit line 22 in `excel_clone.py`:

```python
TEMPLATE_DIR = r"C:\Users\r.cunha\AppData\Local\Temp\xlsx_template"
```

## Usage

1. Open your Excel file normally
2. Run the script:

```bash
python excel_clone.py
```

3. If multiple workbooks are open, you'll be prompted to choose:

```
Open workbooks:
  1. Budget_2026.xlsx (active)
  2. Sales_Report.xlsx
  0. All of them

Which one? [1]:
```

4. Output is saved to: `%LOCALAPPDATA%\Temp\<source_filename>.xlsx`

## What gets preserved

| Feature | Preserved |
|---|---|
| Cell values (text, numbers) | Yes |
| Formatting from template (styles.xml) | Yes |
| Multiple sheets | Yes |
| Formulas | Values only (not the formulas themselves) |
| Charts / images | Only if present in template |

## Output location

All output files are saved to:

```
C:\Users\<username>\AppData\Local\Temp\<source_filename>.xlsx
```
