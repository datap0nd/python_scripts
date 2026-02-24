"""
Excel File Cloner — bypasses application-level file saves.

Approach 1: SaveCopyAs → unzip XML → rezip as new .xlsx (fast, perfect clone)
Approach 2: Read via COM object model → rebuild with openpyxl (fallback if Approach 1 fails)

Output: %LOCALAPPDATA%\Temp\<source_filename>.xlsx
"""

import win32com.client
import zipfile
import os
import shutil
import sys
from pathlib import Path

# ─── Config ───────────────────────────────────────────────────────────────────

TEMP_DIR = os.path.join(os.environ["LOCALAPPDATA"], "Temp")
TEMP_COPY = os.path.join(TEMP_DIR, "_xlclone_temp.xlsx")
TEMP_EXTRACT = os.path.join(TEMP_DIR, "_xlclone_xml")


# ─── Approach 1: SaveCopyAs + rezip ──────────────────────────────────────────

def approach_1(wb, output_path):
    """Clone via SaveCopyAs → unzip → rezip. Perfect 1:1 copy."""
    print("\n[Approach 1] SaveCopyAs + rezip")
    print("  Saving temp copy...")

    wb.SaveCopyAs(TEMP_COPY)

    # Check if the copy is a valid zip (not encrypted)
    try:
        with zipfile.ZipFile(TEMP_COPY, "r") as z:
            z.testzip()
        print("  Temp copy is a valid xlsx (zip). Proceeding...")
    except (zipfile.BadZipFile, Exception) as e:
        print(f"  FAILED — file is not a valid zip: {e}")
        print("  The file is likely encrypted on disk. Falling back to Approach 2.")
        if os.path.exists(TEMP_COPY):
            os.remove(TEMP_COPY)
        return False

    # Extract all XML
    if os.path.exists(TEMP_EXTRACT):
        shutil.rmtree(TEMP_EXTRACT)

    with zipfile.ZipFile(TEMP_COPY, "r") as z:
        z.extractall(TEMP_EXTRACT)
        file_list = z.namelist()

    print(f"  Extracted {len(file_list)} files from xlsx")

    # Rezip into output (pure Python — no Excel involved)
    if os.path.exists(output_path):
        os.remove(output_path)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for root, _, files in os.walk(TEMP_EXTRACT):
            for f in files:
                full_path = os.path.join(root, f)
                arcname = os.path.relpath(full_path, TEMP_EXTRACT)
                zout.write(full_path, arcname)

    # Cleanup
    os.remove(TEMP_COPY)
    shutil.rmtree(TEMP_EXTRACT)

    print(f"  SUCCESS → {output_path}")
    return True


# ─── Approach 2: COM object model + openpyxl rebuild ─────────────────────────

def approach_2(wb, output_path):
    """Read all data + formatting via COM, rebuild from scratch with openpyxl."""
    print("\n[Approach 2] COM read + openpyxl rebuild")

    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    out_wb = Workbook()
    out_wb.remove(out_wb.active)

    total_sheets = wb.Sheets.Count
    for si in range(1, total_sheets + 1):
        ws = wb.Sheets(si)
        out_ws = out_wb.create_sheet(title=ws.Name)
        print(f"\n  Sheet {si}/{total_sheets}: '{ws.Name}'")

        # Get used range
        used = ws.UsedRange
        if used is None:
            print("    (empty)")
            continue

        start_row = used.Row
        start_col = used.Column
        num_rows = used.Rows.Count
        num_cols = used.Columns.Count
        print(f"    Range: {num_rows} rows x {num_cols} cols")

        # ── Bulk read values (much faster than cell-by-cell) ──
        raw = used.Value
        if num_rows == 1 and num_cols == 1:
            values = [[raw]]
        elif num_rows == 1:
            values = [list(raw)]
        elif num_cols == 1:
            values = [[v] for (v,) in raw] if isinstance(raw[0], tuple) else [[v] for v in raw]
        else:
            values = [list(row) for row in raw]

        # ── Write values + copy formatting ──
        for i, row_data in enumerate(values):
            for j, val in enumerate(row_data):
                r = start_row + i
                c = start_col + j
                out_cell = out_ws.cell(row=r, column=c, value=val)

                try:
                    src = ws.Cells(r, c)
                    _copy_font(src, out_cell)
                    _copy_fill(src, out_cell)
                    _copy_alignment(src, out_cell)
                    _copy_number_format(src, out_cell)
                    _copy_borders(src, out_cell)
                except Exception:
                    pass

            # Progress for large sheets
            if num_rows > 100 and (i + 1) % 500 == 0:
                print(f"    ... {i + 1}/{num_rows} rows")

        # ── Column widths ──
        for c in range(start_col, start_col + num_cols):
            try:
                w = ws.Columns(c).ColumnWidth
                if w:
                    out_ws.column_dimensions[get_column_letter(c)].width = float(w)
            except Exception:
                pass

        # ── Row heights ──
        for r in range(start_row, start_row + num_rows):
            try:
                h = ws.Rows(r).RowHeight
                if h:
                    out_ws.row_dimensions[r].height = float(h)
            except Exception:
                pass

        # ── Merged cells ──
        merged_done = set()
        for i in range(num_rows):
            for j in range(num_cols):
                r = start_row + i
                c = start_col + j
                try:
                    cell = ws.Cells(r, c)
                    if cell.MergeCells:
                        addr = cell.MergeArea.Address.replace("$", "")
                        if addr not in merged_done:
                            merged_done.add(addr)
                            out_ws.merge_cells(addr)
                except Exception:
                    pass

        if merged_done:
            print(f"    Merged regions: {len(merged_done)}")

    out_wb.save(output_path)
    print(f"\n  SUCCESS → {output_path}")
    return True


# ─── Formatting helpers ──────────────────────────────────────────────────────

def _bgr_to_hex(color_int):
    """Convert Windows BGR color integer to hex RGB string."""
    c = int(color_int)
    b = (c >> 16) & 0xFF
    g = (c >> 8) & 0xFF
    r = c & 0xFF
    return f"{r:02X}{g:02X}{b:02X}"


def _copy_font(src, out_cell):
    from openpyxl.styles import Font
    sf = src.Font
    kwargs = {}
    if sf.Name:
        kwargs["name"] = sf.Name
    if sf.Size:
        kwargs["size"] = sf.Size
    if sf.Bold is not None:
        kwargs["bold"] = bool(sf.Bold)
    if sf.Italic is not None:
        kwargs["italic"] = bool(sf.Italic)
    if sf.Strikethrough is not None:
        kwargs["strike"] = bool(sf.Strikethrough)
    # Underline: 2 = xlUnderlineStyleSingle, -4142 = xlNone
    if sf.Underline and sf.Underline == 2:
        kwargs["underline"] = "single"
    elif sf.Underline and sf.Underline == 4:
        kwargs["underline"] = "double"
    try:
        if sf.Color and sf.Color.RGB and sf.Color.RGB != 0:
            rgb = sf.Color.RGB
            if isinstance(rgb, int):
                kwargs["color"] = _bgr_to_hex(rgb)
    except Exception:
        pass
    if kwargs:
        out_cell.font = Font(**kwargs)


def _copy_fill(src, out_cell):
    from openpyxl.styles import PatternFill
    interior = src.Interior
    # Pattern: -4142 = xlNone, 1 = xlSolid
    if interior.Pattern and interior.Pattern != -4142:
        try:
            rgb = _bgr_to_hex(interior.Color)
            out_cell.fill = PatternFill(
                start_color=rgb, end_color=rgb, fill_type="solid"
            )
        except Exception:
            pass


def _copy_alignment(src, out_cell):
    from openpyxl.styles import Alignment
    ha_map = {
        -4131: "left", -4108: "center", -4152: "right",
        -4130: "justify", 1: "general", 5: "fill", 7: "distributed",
    }
    va_map = {
        -4160: "top", -4108: "center", -4107: "bottom",
        -4130: "justify", 5: "distributed",
    }
    out_cell.alignment = Alignment(
        horizontal=ha_map.get(src.HorizontalAlignment, "general"),
        vertical=va_map.get(src.VerticalAlignment, "bottom"),
        wrap_text=bool(src.WrapText) if src.WrapText else False,
        text_rotation=int(src.Orientation) if src.Orientation else 0,
        indent=int(src.IndentLevel) if src.IndentLevel else 0,
    )


def _copy_number_format(src, out_cell):
    nf = src.NumberFormat
    if nf:
        out_cell.number_format = str(nf)


def _copy_borders(src, out_cell):
    from openpyxl.styles import Border, Side

    STYLE_MAP = {
        1: "thin", 2: "medium", 3: "dashed", 4: "dotted",
        5: "thick", 6: "double", 7: "hair",
        8: "mediumDashed", 9: "dashDot", 10: "mediumDashDot",
        11: "dashDotDot", 12: "mediumDashDotDot", 13: "slantDashDot",
    }
    EDGE_MAP = {
        "left": 7,    # xlEdgeLeft
        "right": 10,  # xlEdgeRight
        "top": 8,     # xlEdgeTop
        "bottom": 9,  # xlEdgeBottom
    }

    sides = {}
    for name, edge_idx in EDGE_MAP.items():
        try:
            border = src.Borders(edge_idx)
            style = STYLE_MAP.get(border.LineStyle)
            if style:
                kwargs = {"style": style}
                try:
                    if border.Color is not None:
                        kwargs["color"] = _bgr_to_hex(border.Color)
                except Exception:
                    pass
                sides[name] = Side(**kwargs)
        except Exception:
            pass

    if sides:
        out_cell.border = Border(**sides)


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
            active = " ← (active)" if w.Name == excel.ActiveWorkbook.Name else ""
            print(f"  {i}. {w.Name}{active}")
        print(f"  0. All of them")

        choice = input("\nWhich one? [1]: ").strip()
        if choice == "0":
            for w in workbooks:
                _clone(w, os.path.join(TEMP_DIR, w.Name))
            print("\nDone! All files saved.")
            return
        elif choice == "":
            wb = workbooks[0]
        else:
            wb = workbooks[int(choice) - 1]

    source_name = wb.Name
    output_path = os.path.join(TEMP_DIR, source_name)

    _clone(wb, output_path)
    print(f"\nDone! File saved to:\n  {output_path}")


def _clone(wb, output_path):
    """Run Approach 1, fall back to Approach 2."""
    print(f"\nSource:  {wb.Name}")
    print(f"Output:  {output_path}")
    if not approach_1(wb, output_path):
        approach_2(wb, output_path)


if __name__ == "__main__":
    main()
