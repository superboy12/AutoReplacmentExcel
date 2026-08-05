"""
=============================================================================
  inspect_excel_com.py
  Script inspeksi menyeluruh file Excel menggunakan Microsoft Excel COM.
  Membaca: cell values, shapes, textboxes, drawings, dan border styles.

  Jalankan: python inspect_excel_com.py
=============================================================================
"""

import os
import pythoncom
import win32com.client
from datetime import datetime

# ============================================================
# Ubah path ke salah satu file Excel yang ingin diperiksa:
FILE_PATH = r"D:\Excel\contoh.xlsx"
# ============================================================

XL_WHOLE = 1
XL_PART  = 2


def col_index_to_letter(n: int) -> str:
    """Konversi indeks kolom (1-based) ke huruf Excel."""
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def border_style_name(style_val) -> str:
    """Konversi nilai border style ke nama."""
    styles = {
        -4142: "none", 1: "hair", 2: "thin", 3: "medium",
        4: "thick", 5: "double", 6: "dotted", 7: "dash-dot",
        8: "dash-dot-dot", 9: "slant-dash-dot", 13: "medium-dash-dot",
        14: "medium-dash-dot-dot",
    }
    try:
        return styles.get(int(style_val), f"({style_val})")
    except Exception:
        return "?"


def inspect_file(filepath: str):
    pythoncom.CoInitialize()
    excel = None
    wb    = None

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible        = False
        excel.DisplayAlerts  = False
        excel.AskToUpdateLinks = False

        abs_path = os.path.abspath(filepath)
        print(f"\n{'='*72}")
        print(f"  File   : {abs_path}")
        print(f"  Waktu  : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'='*72}")

        wb = excel.Workbooks.Open(abs_path, UpdateLinks=False, ReadOnly=True)
        print(f"\n  Sheet  : {[wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count+1)]}")
        print(f"  Total  : {wb.Worksheets.Count} sheet")

        for sh_idx in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(sh_idx)
            ws_name = ws.Name

            print(f"\n{'─'*72}")
            print(f"  ── SHEET [{ws_name}] ─────────────────────────────────────────")
            print(f"{'─'*72}")
            print(f"  UsedRange: {ws.UsedRange.Address}")

            # ── 1. CELL VALUES ──────────────────────────────────────────────
            print(f"\n  [CELLS - non-kosong]")
            used = ws.UsedRange
            row_start = used.Row
            col_start = used.Column
            row_count = used.Rows.Count
            col_count = used.Columns.Count

            printed_cells = 0
            for r in range(row_start, row_start + row_count):
                for c in range(col_start, col_start + col_count):
                    cell = ws.Cells(r, c)
                    val  = cell.Value
                    if val is None or str(val).strip() == "":
                        continue
                    coord = f"{col_index_to_letter(c)}{r}"

                    # Border info
                    try:
                        borders = cell.Borders
                        b_info = []
                        b_map = {7: "Top", 8: "Bot", 9: "Left", 10: "Right"}
                        for b_const, b_label in b_map.items():
                            b = borders(b_const)
                            s = border_style_name(b.LineStyle)
                            if s != "none":
                                b_info.append(f"{b_label}:{s}")
                        border_str = " | ".join(b_info) if b_info else ""
                    except Exception:
                        border_str = ""

                    merge_info = ""
                    try:
                        if cell.MergeCells:
                            ma = cell.MergeArea.Address
                            merge_info = f" [MERGED:{ma}]"
                    except Exception:
                        pass

                    val_str = str(val)[:80]
                    b_str   = f"  border:[{border_str}]" if border_str else ""
                    print(f"    {coord:>6} │ {val_str}{merge_info}{b_str}")
                    printed_cells += 1
                    if printed_cells >= 60:
                        print(f"    ... (baris selanjutnya tidak ditampilkan)")
                        break
                if printed_cells >= 60:
                    break

            # ── 2. SHAPES (textbox, garis, logo, ttd area) ─────────────────
            print(f"\n  [SHAPES ({ws.Shapes.Count} total)]")
            if ws.Shapes.Count == 0:
                print("    (tidak ada shape)")
            else:
                for sh_i in range(1, ws.Shapes.Count + 1):
                    shape = ws.Shapes(sh_i)
                    try:
                        sh_name = shape.Name
                        sh_type = shape.Type
                        sh_left = round(shape.Left, 1)
                        sh_top  = round(shape.Top, 1)
                        sh_w    = round(shape.Width, 1)
                        sh_h    = round(shape.Height, 1)

                        # Coba baca teks dalam shape
                        try:
                            text_val = shape.TextFrame.Characters().Text
                            text_val = text_val.replace("\n", "↵").replace("\r", "")[:80]
                        except Exception:
                            text_val = "(no text)"

                        # Coba baca line (garis)
                        try:
                            line_style  = shape.Line.Style
                            line_weight = round(shape.Line.Weight, 2)
                            line_color  = shape.Line.ForeColor.RGB
                            line_info   = f" LINE weight={line_weight} color=#{line_color:06X}"
                        except Exception:
                            line_info = ""

                        print(f"    [{sh_i:>2}] {sh_name:<25} type={sh_type} "
                              f"pos=({sh_left},{sh_top}) size=({sh_w}x{sh_h})")
                        if text_val != "(no text)":
                            print(f"         TEXT: \"{text_val}\"")
                        if line_info:
                            print(f"         {line_info}")
                    except Exception as e:
                        print(f"    [{sh_i:>2}] (error membaca shape: {e})")

        wb.Close(SaveChanges=False)
        wb = None

    except Exception as e:
        print(f"\n  [ERROR] {e}")
    finally:
        if wb:
            try: wb.Close(SaveChanges=False)
            except: pass
        if excel:
            try: excel.Quit()
            except: pass
        pythoncom.CoUninitialize()


# ── Entry Point ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if not os.path.isfile(FILE_PATH):
        raw = input(f"  File tidak ditemukan. Masukkan path file Excel: ").strip().strip('"')
        FILE_PATH_ACTUAL = raw
    else:
        FILE_PATH_ACTUAL = FILE_PATH

    inspect_file(FILE_PATH_ACTUAL)
