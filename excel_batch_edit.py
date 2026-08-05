"""
=============================================================================
  Excel Batch Editor - excel_batch_edit.py
  Melakukan Find & Replace massal pada semua file .xlsx dalam satu folder.

  Engine  : Microsoft Excel COM (pywin32)
  Install : pip install pywin32
  Jalankan: python excel_batch_edit.py

  Keunggulan dibanding openpyxl:
  - Excel sendiri yang memproses file → 0% risiko korupsi
  - Semua drawing, flowchart, logo, shape tetap 100% utuh
  - Format, rumus, merge cell, warna tidak pernah berubah
=============================================================================
"""

import os
import re
import time
import shutil
import pythoncom
import win32com.client
from datetime import datetime


# =============================================================================
# =================== UBAH BAGIAN INI SESUAI KEBUTUHAN ====================

# Folder default. Biarkan kosong agar selalu ditanya saat dijalankan.
# Atau isi langsung: r"D:\Data\SOP Excel"
FOLDER_INPUT = r""

# Mode output:
#   "overwrite"  → Timpa file asli (direkomendasikan, ada backup otomatis)
#   "new_folder" → Simpan ke FOLDER_OUTPUT
OUTPUT_MODE = "overwrite"

# Folder tujuan jika OUTPUT_MODE = "new_folder"
FOLDER_OUTPUT = r"C:\Excel\Output"

# Buat backup otomatis sebelum menimpa file asli?
# Backup disimpan di subfolder "_backup" di dalam FOLDER_INPUT
BACKUP_ENABLED = True

# Daftar aturan Find & Replace
# Format: "Teks Yang Dicari" : "Teks Pengganti"
# Gunakan string kosong "" untuk MENGHAPUS teks tertentu.
REPLACEMENTS = [
    # Jabatan TTD: tambahkan Plt
    # (Berlaku di semua sheet, ditandai dengan sheet: "*")
    {
        "find": "KEPALA DINAS KEHUTANAN",
        "replace": "Plt KEPALA DINAS KEHUTANAN",
        "sheet": "*"
    },
    {
        "find": "Kepala Dinas Kehutanan Provinsi Lampung",
        "replace": "Plt Kepala Dinas Kehutanan Provinsi Lampung",
        "sheet": "*"
    },

    # Nama penanda tangan 
    # (HANYA di sheet ke-3, agar TIDAK merusak "Kualifikasi Pelaksana" di sheet 2)
    {
        "find": "Ir. YANYAN RUCHYANSYAH, M.Si.",
        "replace": "Dr. Drs. SULPAKAR, M.M.",
        "sheet": 3
    },
    {
        "find": "YANYAN RUCHYANSYAH, M.Si.",
        "replace": "Dr. Drs. SULPAKAR, M.M.",
        "sheet": 3
    },
    {
        "find": "Ir. YANYAN RUCHYANSYAH",
        "replace": "Dr. Drs. SULPAKAR, M.M.",
        "sheet": 3
    }
]


# Aktifkan pencarian Case-Sensitive?
# True  → "nama" ≠ "Nama"  (lebih presisi)
# False → "nama" == "Nama" (lebih luas)
CASE_SENSITIVE = True

# Proses juga file di subfolder secara rekursif?
INCLUDE_SUBFOLDERS = True

# ---------------------------------------------------------------------------
# TULIS NILAI KE CELL TERTENTU (menulis nilai ke cell spesifik di sheet tertentu)
#
# Mode 1 – Cari label, isi kolom nilai di baris yang sama:
#   { "sheet": "identitas", "find_label": "Tanggal Revisi",
#     "value_col": "L", "value": "13", "only_if_empty": False }
#
# Mode 2 – Langsung via koordinat cell:
#   { "sheet": "identitas", "cell": "L5", "value": "13", "only_if_empty": False }
#
# Kosongkan list [] jika tidak diperlukan.
CELL_WRITES = [
    # ── Tanggal Revisi (split 3 cell: tanggal/bulan/tahun) ─────────────────
    {
        "sheet"        : "identitas",
        "find_label"   : "Tanggal Revisi",
        "value_col"    : "L",
        "value"        : "13",
        "only_if_empty": False,
    },
    {
        "sheet"        : "identitas",
        "find_label"   : "Tanggal Revisi",
        "value_col"    : "M",
        "value"        : "Juni",
        "only_if_empty": False,
    },
    {
        "sheet"        : "identitas",
        "find_label"   : "Tanggal Revisi",
        "value_col"    : "N",
        "value"        : "2026",
        "only_if_empty": False,
    },

    # ── Halaman TTD: isi nama, pangkat, NIP di sheet 3 ─────────────────
    # Mode 3 (find_ref + row_offset):
    #   Cari cell yang berisi teks referensi (jabatan),
    #   lalu tulis nilai ke baris N di bawahnya (row_offset).
    #
    #   "sheet": 3 = Hanya berlaku di sheet ke-3.
    #
    # ── Nama ────────────────────────────────────────────────────────────────
    {
        "sheet"        : 3,                            # hanya sheet 3
        "find_ref"     : "Plt KEPALA DINAS KEHUTANAN", # cari cell jabatan
        "row_offset"   : 6,                            # 6 baris di bawah jabatan
        "value"        : "Dr. Drs. SULPAKAR, M.M.",
        "only_if_empty": True,    # True = hanya isi jika sel masih kosong
    },
    # ── Pangkat ─────────────────────────────────────────────────────────────
    {
        "sheet"        : 3,
        "find_ref"     : "Plt KEPALA DINAS KEHUTANAN",
        "row_offset"   : 7,
        "value"        : "Pembina Utama Madya/IV d",
        "only_if_empty": True,
    },
    # ── NIP ─────────────────────────────────────────────────────────────────
    {
        "sheet"        : 3,
        "find_ref"     : "Plt KEPALA DINAS KEHUTANAN",
        "row_offset"   : 8,
        "value"        : "NIP. 19690205 198910 1 002",
        "only_if_empty": True,
    },
]
# ---------------------------------------------------------------------------

# =================== AKHIR BAGIAN KONFIGURASI =============================
# =============================================================================


# ── Konstanta Excel COM ────────────────────────────────────────────────────
XL_PART  = 2    # LookAt: cari teks yang mengandung string (bukan exact match)
XL_WHOLE = 1    # LookAt: exact match
XL_BY_ROWS = 1  # SearchOrder: baris per baris


def log(message: str, level: str = "INFO"):
    """Cetak pesan log ke konsol dengan timestamp."""
    icons = {"INFO": "ℹ", "OK": "✔", "WARN": "⚠", "ERROR": "✖", "SKIP": "↷"}
    icon  = icons.get(level, "•")
    ts    = datetime.now().strftime("%H:%M:%S")
    print(f"  [{ts}] {icon}  {message}")


def print_separator(char: str = "─", width: int = 70):
    print(f"  {char * width}")


def get_excel_files(folder: str, recursive: bool) -> list[str]:
    """Kumpulkan semua file .xlsx dari folder (dan subfolder jika rekursif)."""
    files = []
    if recursive:
        for root, _, filenames in os.walk(folder):
            if "_backup" in root or "_output" in root:
                continue
            for fname in filenames:
                if fname.lower().endswith(".xlsx") and not fname.startswith("~$"):
                    files.append(os.path.join(root, fname))
    else:
        for fname in os.listdir(folder):
            if fname.lower().endswith(".xlsx") and not fname.startswith("~$"):
                files.append(os.path.join(folder, fname))
    return sorted(files)


def make_backup(filepath: str, backup_dir: str):
    """Buat salinan backup sebelum file dimodifikasi."""
    os.makedirs(backup_dir, exist_ok=True)
    rel_name = os.path.basename(filepath)
    dest = os.path.join(backup_dir, rel_name)
    if os.path.exists(dest):
        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
        name, ext = os.path.splitext(rel_name)
        dest = os.path.join(backup_dir, f"{name}_{ts}{ext}")
    shutil.copy2(filepath, dest)


def col_letter_to_index(col_letter: str) -> int:
    """Konversi huruf kolom Excel ke indeks numerik (A=1, B=2, ...)."""
    col_letter = col_letter.strip().upper()
    result = 0
    for ch in col_letter:
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result


def apply_cell_writes_com(ws_com, cell_writes: list, sheet_name: str) -> int:
    """
    Tulis nilai ke cell tertentu di satu worksheet via COM.

    Mendukung 3 mode per entri:
      Mode 1 – find_label : cari cell berisi label, tulis di kolom target baris sama
      Mode 2 – cell       : tulis langsung ke koordinat (mis. "L5")
      Mode 3 – find_ref   : cari cell referensi, tulis di baris N di bawahnya
                             (berguna untuk mengisi cell kosong di area TTD)

    Mengembalikan jumlah cell yang berhasil ditulis.
    """
    total = 0
    for entry in cell_writes:
        target_sheet  = entry.get("sheet", "*")
        new_value     = entry.get("value", "")
        only_if_empty = entry.get("only_if_empty", False)

        # Cek kecocokan sheet: "*" = semua sheet, int = index sheet, str = nama sheet
        if target_sheet != "*":
            if isinstance(target_sheet, int):
                if ws_com.Index != target_sheet:
                    continue
            elif str(target_sheet).strip().lower() != sheet_name.strip().lower():
                continue

        try:
            # ── Mode 1: find_label ─────────────────────────────────────────
            if "find_label" in entry:
                label = entry["find_label"]
                found = ws_com.Cells.Find(
                    What=label, LookAt=XL_WHOLE,
                    MatchCase=False, SearchFormat=False)
                if found is None:
                    found = ws_com.Cells.Find(
                        What=label, LookAt=XL_PART,
                        MatchCase=False, SearchFormat=False)
                if found is None:
                    continue
                value_col = entry.get("value_col", "L")
                col_idx   = col_letter_to_index(value_col)
                target    = ws_com.Cells(found.Row, col_idx)

            # ── Mode 3: find_ref + row_offset ──────────────────────────────
            elif "find_ref" in entry:
                ref_text   = entry["find_ref"]
                row_offset = entry.get("row_offset", 5)

                # Cari cell referensi (jabatan)
                found_ref = ws_com.Cells.Find(
                    What=ref_text, LookAt=XL_PART,
                    MatchCase=False, SearchFormat=False)
                if found_ref is None:
                    continue

                # Hitung target: baris = ref + offset, kolom = sama dengan ref
                target_row = found_ref.Row + row_offset
                target_col = found_ref.Column
                if "value_col" in entry:
                    target_col = col_letter_to_index(entry["value_col"])
                target = ws_com.Cells(target_row, target_col)

            # ── Mode 2: cell langsung ──────────────────────────────────────
            elif "cell" in entry:
                target = ws_com.Range(entry["cell"])
            else:
                continue

            if only_if_empty and target.Value not in (None, ""):
                continue

            target.Value = new_value
            total += 1

        except Exception as e:
            log(f"CELL_WRITE gagal ({entry.get('value','?')}): {e}", "WARN")

    return total


def smart_cell_replace(ws, find_text: str, replace_text: str,
                       case_sensitive: bool) -> int:
    """
    Ganti find_text → replace_text di sheet ws, dengan cek idempoten per-cell:
      - Jika cell SUDAH mengandung replace_text, cell tersebut di-SKIP.
      - Ini mencegah "Plt Plt ..." atau duplikasi apapun, bahkan jika script
        dijalankan berulang kali pada file yang sudah diproses.

    Mengembalikan jumlah cell yang berhasil diganti.
    """
    try:
        found = ws.Cells.Find(
            What=find_text,
            LookAt=XL_PART,
            MatchCase=case_sensitive,
            SearchFormat=False,
        )
        if found is None:
            return 0

        first_addr = found.Address
        count = 0

        while True:
            cell_val = str(found.Value) if found.Value is not None else ""

            # ── Cek idempoten: skip jika cell SUDAH dalam keadaan benar ─────────
            #
            # Masalah: "Plt. KEPALA DINAS KEHUTANAN" (titik) ≠ "Plt KEPALA..."
            # sehingga cek biasa gagal → muncul "Plt. Plt KEPALA...".
            #
            # Solusi: jika aturan bertujuan menambahkan prefix "Plt",
            # gunakan regex yang mengenali SEMUA varian:
            #   "Plt KEPALA..."   → sudah benar, skip
            #   "Plt. KEPALA..."  → sudah benar, skip
            #   "PLT KEPALA..."   → sudah benar, skip
            # ────────────────────────────────────────────────────────────────────
            if replace_text.lower().startswith("plt "):
                # Deteksi varian Plt: opsional titik, opsional spasi ganda
                guard_pat = r'(?i)Plt\.?\s+' + re.escape(find_text)
                already_correct = bool(re.search(guard_pat, cell_val))
            elif case_sensitive:
                already_correct = replace_text in cell_val
            else:
                already_correct = replace_text.lower() in cell_val.lower()

            if not already_correct:
                if case_sensitive:
                    new_val = cell_val.replace(find_text, replace_text)
                else:
                    new_val = re.sub(
                        re.escape(find_text), replace_text, cell_val,
                        flags=re.IGNORECASE,
                    )
                found.Value = new_val
                count += 1


            nxt = ws.Cells.FindNext(found)
            if nxt is None or nxt.Address == first_addr:
                break
            found = nxt

        return count

    except Exception:
        return 0


def process_file_com(excel_app, filepath: str, replacements: dict,
                     case_sensitive: bool, cell_writes: list,
                     output_path: str) -> dict:
    """
    Proses satu file Excel via COM:
      - Buka file dengan Excel
      - Find & Replace di semua sheet
      - Tulis cell tertentu (CELL_WRITES)
      - Simpan

    Excel menangani semuanya → format/drawing/shape 100% aman.
    """
    result = {"success": False, "changes": 0, "changed_cells": [], "error": ""}
    wb = None

    try:
        abs_filepath = os.path.abspath(filepath)

        # Buka workbook (UpdateLinks=False agar tidak ada prompt)
        wb = excel_app.Workbooks.Open(
            abs_filepath,
            UpdateLinks=False,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
        )

        total_changes = 0

        # ── Find & Replace (smart: skip jika sudah benar, filter per-sheet) ──
        for rule in replacements:
            find_text = rule.get("find", "")
            replace_text = rule.get("replace", "")
            target_sheet = rule.get("sheet", "*")

            if not find_text:
                continue
            
            for ws in wb.Worksheets:
                # Cek filter sheet
                if target_sheet != "*":
                    if isinstance(target_sheet, int):
                        if ws.Index != target_sheet:
                            continue
                    elif str(target_sheet).strip().lower() != ws.Name.lower():
                        continue

                changed = smart_cell_replace(ws, find_text, replace_text, case_sensitive)
                total_changes += changed

        # ── CELL_WRITES (tulis nilai ke cell spesifik) ─────────────────────
        if cell_writes:
            for ws in wb.Worksheets:
                ws_name = ws.Name
                cw_done = apply_cell_writes_com(ws, cell_writes, ws_name)
                total_changes += cw_done

        # ── Simpan ─────────────────────────────────────────────────────────
        abs_output = os.path.abspath(output_path)
        if abs_output == abs_filepath:
            wb.Save()
        else:
            out_dir = os.path.dirname(abs_output)
            if out_dir:
                os.makedirs(out_dir, exist_ok=True)
            # 51 = xlOpenXMLWorkbook (.xlsx)
            wb.SaveAs(abs_output, FileFormat=51)

        wb.Close(SaveChanges=False)
        wb = None

        result["success"] = True
        result["changes"] = total_changes

    except Exception as e:
        err_str = str(e)
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        result["error"] = err_str

    return result


def run_batch():
    """Fungsi utama: jalankan proses batch untuk seluruh file."""

    # ── Kumpulkan file ───────────────────────────────────────────────────────
    excel_files = get_excel_files(FOLDER_INPUT, INCLUDE_SUBFOLDERS)
    total_files = len(excel_files)

    if total_files == 0:
        print(f"\n  [INFO] Tidak ada file .xlsx di folder: {FOLDER_INPUT}")
        return

    print(f"  Folder Input  : {FOLDER_INPUT}")
    print(f"  Mode Output   : {OUTPUT_MODE}")
    print(f"  Backup        : {'Ya' if BACKUP_ENABLED else 'Tidak'}")
    print(f"  Case Sensitive: {'Ya' if CASE_SENSITIVE else 'Tidak'}")
    print(f"  Subfolder     : {'Ya' if INCLUDE_SUBFOLDERS else 'Tidak'}")
    print()
    print(f"  Aturan Find & Replace ({len(REPLACEMENTS)} aturan):")
    for rule in REPLACEMENTS:
        f = rule.get("find", "")
        r = rule.get("replace", "")
        sh = rule.get("sheet", "*")
        r_display = f'"{r}"' if r else "(hapus)"
        print(f'    → [Sheet: {sh}] "{f}"  ⟶  {r_display}')
    print()
    print_separator("═")
    print(f"  Ditemukan {total_files} file Excel. Memulai proses...")
    print_separator("═")

    backup_dir = os.path.join(FOLDER_INPUT, "_backup")

    stats = {
        "total"     : total_files,
        "success"   : 0,
        "skipped"   : 0,
        "failed"    : 0,
        "no_change" : 0,
        "changes"   : 0,
    }
    failed_files  = []
    skipped_files = []
    start_time    = time.time()

    # ── Inisialisasi Excel satu kali untuk semua file ──────────────────────
    pythoncom.CoInitialize()
    excel_app = None

    try:
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible         = False   # Jalankan di background
        excel_app.DisplayAlerts   = False   # Tidak ada dialog/popup
        excel_app.ScreenUpdating  = False   # Lebih cepat
        excel_app.AskToUpdateLinks = False  # Tidak tanya update link

        # ── Proses setiap file ────────────────────────────────────────────
        for idx, filepath in enumerate(excel_files, 1):
            rel_name = os.path.relpath(filepath, FOLDER_INPUT)
            progress = f"[{idx:>3}/{total_files}]"
            print()
            print(f"  {progress} 📄 {rel_name}")

            # Tentukan path output
            if OUTPUT_MODE == "overwrite":
                output_path = filepath
            else:
                rel = os.path.relpath(filepath, FOLDER_INPUT)
                output_path = os.path.join(FOLDER_OUTPUT, rel)

            # Backup sebelum timpa
            if BACKUP_ENABLED and OUTPUT_MODE == "overwrite":
                try:
                    make_backup(filepath, backup_dir)
                except Exception as e:
                    log(f"Backup gagal: {e}", "WARN")

            # Proses via COM
            result = process_file_com(
                excel_app, filepath, REPLACEMENTS,
                CASE_SENSITIVE, CELL_WRITES, output_path
            )

            if not result["success"]:
                err = result["error"]
                # Deteksi file terkunci (Excel sedang membukanya)
                if any(k in err.lower() for k in
                       ["permission", "access", "locked", "being used", "dibuka"]):
                    log(f"DILEWATI → File sedang digunakan/terkunci.", "SKIP")
                    skipped_files.append((rel_name, err))
                    stats["skipped"] += 1
                else:
                    log(f"GAGAL    → {err}", "ERROR")
                    failed_files.append((rel_name, err))
                    stats["failed"] += 1
            else:
                n = result["changes"]
                if n == 0:
                    log("Tidak ada perubahan.", "INFO")
                    stats["no_change"] += 1
                else:
                    log(f"Berhasil → {n} aturan diterapkan.", "OK")
                stats["success"] += 1
                stats["changes"] += n

    finally:
        if excel_app is not None:
            try:
                excel_app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    # ── Laporan Akhir ──────────────────────────────────────────────────────
    elapsed = time.time() - start_time
    print()
    print_separator("═")
    print("  📊  LAPORAN AKHIR")
    print_separator("─")
    print(f"  Total file ditemukan  : {stats['total']}")
    print(f"  ✔  Berhasil diproses  : {stats['success']}")
    print(f"     - Ada perubahan    : {stats['success'] - stats['no_change']}")
    print(f"     - Tanpa perubahan  : {stats['no_change']}")
    print(f"  ↷  Dilewati (terkunci): {stats['skipped']}")
    print(f"  ✖  Gagal              : {stats['failed']}")
    print(f"  ✎  Total aturan dijalankan: {stats['changes']}")
    print(f"  ⏱  Waktu proses       : {elapsed:.2f} detik")
    if BACKUP_ENABLED and OUTPUT_MODE == "overwrite":
        print(f"  💾  Backup tersimpan di : {backup_dir}")

    if skipped_files:
        print()
        print_separator("─")
        print("  ↷  File yang DILEWATI (sedang dibuka/terkunci):")
        for fname, _ in skipped_files:
            print(f"       • {fname}")

    if failed_files:
        print()
        print_separator("─")
        print("  ✖  File yang GAGAL:")
        for fname, reason in failed_files:
            print(f"       • {fname}")
            print(f"         Error: {reason[:120]}")

    print_separator("═")
    print()


def ask_user() -> bool:
    """Menu interaktif saat script dijalankan."""
    global FOLDER_INPUT, OUTPUT_MODE, FOLDER_OUTPUT, BACKUP_ENABLED

    print()
    print("  ╔══════════════════════════════════════════════════════════════╗")
    print("  ║     EXCEL BATCH EDITOR  •  Microsoft Excel COM engine      ║")
    print("  ╚══════════════════════════════════════════════════════════════╝")
    print()
    print("  Aturan Find & Replace yang aktif:")
    for rule in REPLACEMENTS:
        f = rule.get("find", "")
        r = rule.get("replace", "")
        sh = rule.get("sheet", "*")
        r_disp = f'"{r}"' if r else "(hapus)"
        print(f'    • [Sheet: {sh}] "{f}"  →  {r_disp}')
    if CELL_WRITES:
        print("  Penulisan Cell langsung:")
        for cw in CELL_WRITES:
            label = cw.get("find_label") or cw.get("cell", "-")
            print(f"    • [{cw.get('sheet')}] {label}  →  \"{cw.get('value')}\"")
    print()
    print_separator("─")

    # ── 1. Folder Input ──────────────────────────────────────────────────
    if FOLDER_INPUT and os.path.isdir(FOLDER_INPUT):
        print(f"  Folder default  : {FOLDER_INPUT}")
        pilihan = input("  Pakai folder ini? [Y/n] : ").strip().lower()
        if pilihan in ("n", "tidak", "no"):
            FOLDER_INPUT = ""
    else:
        FOLDER_INPUT = ""

    while not FOLDER_INPUT:
        raw = input("  Masukkan path folder Excel : ").strip().strip('"').strip("'")
        if not raw:
            continue
        if os.path.isdir(raw):
            FOLDER_INPUT = raw
        else:
            print(f"  [!] Folder tidak ditemukan: {raw}")
            print( "      Coba lagi — pastikan path sudah benar.")

    # ── 2. Mode Output ───────────────────────────────────────────────────
    print()
    print("  Mode penyimpanan:")
    print("    [1] Timpa file asli + backup otomatis  ← Default")
    print("    [2] Simpan ke folder output baru")
    mode_input = input("  Pilih [1/2] : ").strip()

    if mode_input == "2":
        OUTPUT_MODE   = "new_folder"
        default_out   = os.path.join(FOLDER_INPUT, "_output")
        raw_out       = input(f"  Folder output [{default_out}] : ").strip().strip('"').strip("'")
        FOLDER_OUTPUT  = raw_out if raw_out else default_out
        BACKUP_ENABLED = False
        print(f"  → Hasil disimpan ke: {FOLDER_OUTPUT}")
    else:
        OUTPUT_MODE    = "overwrite"
        bak            = input("  Buat backup otomatis? [Y/n] : ").strip().lower()
        BACKUP_ENABLED = bak not in ("n", "tidak", "no")
        print(f"  → Timpa file asli. Backup: {'Ya' if BACKUP_ENABLED else 'Tidak'}")

    # ── 3. Konfirmasi ────────────────────────────────────────────────────
    print()
    print_separator("─")
    print(f"  Folder  : {FOLDER_INPUT}")
    print(f"  Output  : {OUTPUT_MODE}" +
          (f" → {FOLDER_OUTPUT}" if OUTPUT_MODE == "new_folder" else ""))
    print(f"  Backup  : {'Ya' if BACKUP_ENABLED else 'Tidak'}")
    print()
    ok = input("  Lanjutkan proses? [Y/n] : ").strip().lower()
    if ok in ("n", "tidak", "no"):
        print("  Dibatalkan.")
        return False
    return True


# ── Entry Point ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if ask_user():
        run_batch()
