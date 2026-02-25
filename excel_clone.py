"""
Excel File Cloner — bypasses application-level file saves.

Duplicates an unzipped xlsx template → populates XMLs with data read
from the open Excel workbook via COM (from memory) → rezips as new .xlsx.

Output: %LOCALAPPDATA%\Temp\<source_filename>.xlsx
"""

import win32com.client
import zipfile
import os
import shutil
import sys
import xml.etree.ElementTree as ET

# ─── Config ───────────────────────────────────────────────────────────────────

TEMP_DIR = os.path.join(os.environ["LOCALAPPDATA"], "Temp")

# *** SET THIS to the path of your unzipped sample xlsx folder ***
TEMPLATE_DIR = r"C:\Users\<your_user>\AppData\Local\Temp\xlsx_template"

WORK_DIR = os.path.join(TEMP_DIR, "_xlclone_work")

# ─── XML setup ────────────────────────────────────────────────────────────────

SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", SPREADSHEET_NS)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")


# ─── Clone logic ─────────────────────────────────────────────────────────────

def clone(wb, output_path):
    """Duplicate template, populate with COM data, rezip."""
    print(f"\nSource:  {wb.Name}")
    print(f"Output:  {output_path}")

    if not os.path.isdir(TEMPLATE_DIR):
        print(f"\nERROR: Template folder not found:\n  {TEMPLATE_DIR}")
        print("\nTo set up the template:")
        print("  1. Take any .xlsx file")
        print("  2. Rename to .zip")
        print("  3. Extract to the folder above")
        return False

    # Duplicate template folder
    if os.path.exists(WORK_DIR):
        shutil.rmtree(WORK_DIR)
    shutil.copytree(TEMPLATE_DIR, WORK_DIR)
    print(f"  Duplicated template")

    # Read data from open Excel via COM (from memory, unencrypted)
    for si in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(si)
        used = ws.UsedRange
        if used is None:
            continue

        start_row = used.Row
        start_col = used.Column
        num_rows = used.Rows.Count
        num_cols = used.Columns.Count
        print(f"  Sheet '{ws.Name}': {num_rows} rows x {num_cols} cols")

        # Bulk read values
        raw = used.Value
        if num_rows == 1 and num_cols == 1:
            values = [[raw]]
        elif num_rows == 1:
            values = [list(raw)]
        elif num_cols == 1:
            values = [[v] for (v,) in raw] if isinstance(raw[0], tuple) else [[v] for v in raw]
        else:
            values = [list(row) for row in raw]

        # Build sheet XML
        sheet_xml = _build_sheet_xml(values, start_row, start_col)

        # Write into the template's worksheet file
        sheet_path = os.path.join(WORK_DIR, "xl", "worksheets", f"sheet{si}.xml")
        with open(sheet_path, "w", encoding="utf-8") as f:
            f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
            f.write(sheet_xml)
        print(f"    Wrote sheet{si}.xml")

    # Build shared strings (empty — we use inline strings)
    _build_shared_strings(WORK_DIR)

    # Zip into output
    if os.path.exists(output_path):
        os.remove(output_path)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for root, _, files in os.walk(WORK_DIR):
            for f in files:
                full_path = os.path.join(root, f)
                arcname = os.path.relpath(full_path, WORK_DIR)
                zout.write(full_path, arcname)

    # Cleanup work dir
    shutil.rmtree(WORK_DIR)

    print(f"\n  SUCCESS → {output_path}")
    return True


# ─── XML builders ────────────────────────────────────────────────────────────

def _build_sheet_xml(values, start_row, start_col):
    """Build worksheet XML from COM-read values."""
    ns = SPREADSHEET_NS
    root = ET.Element(f"{{{ns}}}worksheet")
    sheet_data = ET.SubElement(root, f"{{{ns}}}sheetData")

    for i, row_vals in enumerate(values):
        r = start_row + i
        row_el = ET.SubElement(sheet_data, f"{{{ns}}}row", r=str(r))

        for j, val in enumerate(row_vals):
            c = start_col + j
            col_letter = _col_letter(c)
            cell_ref = f"{col_letter}{r}"
            cell_el = ET.SubElement(row_el, f"{{{ns}}}c", r=cell_ref)

            if val is None:
                continue
            elif isinstance(val, str):
                cell_el.set("t", "inlineStr")
                is_el = ET.SubElement(cell_el, f"{{{ns}}}is")
                t_el = ET.SubElement(is_el, f"{{{ns}}}t")
                t_el.text = val
            elif isinstance(val, bool):
                cell_el.set("t", "b")
                v_el = ET.SubElement(cell_el, f"{{{ns}}}v")
                v_el.text = "1" if val else "0"
            else:
                v_el = ET.SubElement(cell_el, f"{{{ns}}}v")
                v_el.text = str(val)

    return ET.tostring(root, encoding="unicode")


def _build_shared_strings(work_dir):
    """Build empty sharedStrings.xml (we use inline strings instead)."""
    ns = SPREADSHEET_NS
    root = ET.Element(f"{{{ns}}}sst", count="0", uniqueCount="0")
    sst_path = os.path.join(work_dir, "xl", "sharedStrings.xml")
    tree = ET.ElementTree(root)
    with open(sst_path, "wb") as f:
        tree.write(f, xml_declaration=True, encoding="UTF-8")


def _col_letter(col_num):
    """Convert 1-based column number to Excel letter (1→A, 27→AA)."""
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    # Connect to running Excel
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception as e:
        print(f"ERROR: Could not connect to Excel. Is it running?\n  {e}")
        sys.exit(1)

    count = excel.Workbooks.Count
    if count == 0:
        print("ERROR: No workbooks are open in Excel.")
        sys.exit(1)

    # List all open workbooks and let user pick
    workbooks = []
    for i in range(1, count + 1):
        workbooks.append(excel.Workbooks(i))

    if count == 1:
        wb = workbooks[0]
    else:
        print("\nOpen workbooks:")
        for i, w in enumerate(workbooks, 1):
            active = " (active)" if w.Name == excel.ActiveWorkbook.Name else ""
            print(f"  {i}. {w.Name}{active}")
        print(f"  0. All of them")

        choice = input("\nWhich one? [1]: ").strip()
        if choice == "0":
            for w in workbooks:
                clone(w, os.path.join(TEMP_DIR, w.Name))
            print("\nDone! All files saved.")
            return
        elif choice == "":
            wb = workbooks[0]
        else:
            wb = workbooks[int(choice) - 1]

    source_name = wb.Name
    output_path = os.path.join(TEMP_DIR, source_name)

    clone(wb, output_path)
    print(f"\nDone! File saved to:\n  {output_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
    input("\nPress Enter to exit...")
