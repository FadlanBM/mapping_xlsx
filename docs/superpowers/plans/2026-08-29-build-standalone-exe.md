# Build Standalone Windows EXE Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build and package the `rekap-transaksi` application into a standalone Windows `.exe` executable that can run on any Windows laptop without needing Python installed, supporting both interactive double-click (GUI file dialog / interactive menu) and CLI command-line arguments.

**Architecture:** Enhance `main.py` to handle both CLI arguments and interactive double-click / drag-and-drop mode (with `input()` pause and fallback file dialog), create a clean PyInstaller build configuration script (`build_exe.py` and `RekapTransaksi.spec`), and package everything into a single standalone `.exe` (`dist/RekapTransaksi.exe`).

**Tech Stack:** Python 3.14, PyInstaller 6.22, pandas, openpyxl, tkinter (standard library for file picker dialog), unittest.

## Global Constraints

- Standalone executable must be self-contained in `dist/` and run without requiring Python or pip on target machines.
- Preserve 100% backward compatibility with existing CLI arguments (`-i`, `-o`, `--generate-sample`, `--order`, `--add-totals`).
- If double-clicked without arguments, display an interactive menu / file dialog and pause before closing to prevent the window from abruptly vanishing.
- Fully tested with automated verification steps.

---

### Task 1: Enhance `main.py` with Interactive Mode & Drag-and-Drop Support

**Files:**
- Modify: `main.py`
- Test: `tests/test_cli_interactive.py`

**Interfaces:**
- Consumes: `detect_and_load_data`, `sort_by_bukti_and_type`, `export_to_excel`, `generate_sample_excel` from `rekap_transaksi.py`
- Produces: `run_interactive_menu()`, `process_file(input_path, output_path, order, add_totals)`, `main()` entrypoint

- [ ] **Step 1: Write the failing test for interactive/helper functions**

```python
# tests/test_cli_interactive.py
import unittest
import os
import sys
from unittest.mock import patch
import pandas as pd
from main import process_file

class TestCLIInteractive(unittest.TestCase):
    def setUp(self):
        self.sample_input = "tests/test_sample_cli.xlsx"
        self.sample_output = "tests/test_output_cli.xlsx"
        df = pd.DataFrame({
            "BUKTI": ["B001", "B001"],
            "NAMA AKUN": ["Kas", "Pendapatan"],
            "DEBET": [100000, 0],
            "KREDIT": [0, 100000],
            "KETERANGAN": ["Penerimaan", "Penerimaan"]
        })
        df.to_excel(self.sample_input, index=False)

    def tearDown(self):
        if os.path.exists(self.sample_input):
            os.remove(self.sample_input)
        if os.path.exists(self.sample_output):
            os.remove(self.sample_output)

    def test_process_file_success(self):
        success = process_file(self.sample_input, self.sample_output, order="debet-first", add_totals=True)
        self.assertTrue(success)
        self.assertTrue(os.path.exists(self.sample_output))

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests/test_cli_interactive.py`
Expected: FAIL with `ImportError: cannot import name 'process_file' from 'main'`

- [ ] **Step 3: Update `main.py` to add `process_file`, interactive mode, and drag-and-drop handling**

Update `main.py` with:
```python
import argparse
import os
import sys
import pandas as pd
from rekap_transaksi import (
    detect_and_load_data, 
    sort_by_bukti_and_type, 
    export_to_excel, 
    generate_sample_excel
)

def process_file(input_path: str, output_path: str, order: str = "debet-first", add_totals: bool = False) -> bool:
    """Proses file Excel/CSV rekap transaksi."""
    if not os.path.exists(input_path):
        print(f"[ERROR] File input '{input_path}' tidak ditemukan!")
        return False

    print(f"\n[1/3] Membaca & mendeteksi header file: {input_path}...")
    try:
        df = detect_and_load_data(input_path)
    except Exception as e:
        print(f"[ERROR] Gagal membaca file: {e}")
        return False

    print(f"[2/3] Memetakan & mengurutkan {len(df)} baris data...")
    debet_first = (order == "debet-first")
    df_sorted = sort_by_bukti_and_type(df, debet_first=debet_first)

    print(f"[3/3] Mengekspor {len(df_sorted)} baris data ke: {output_path}...")
    try:
        saved_path = export_to_excel(df_sorted, output_path, add_totals=add_totals)
        print(f"[SUKSES] Data berhasil di-mapping dan diekspor ke: {saved_path}")
        return True
    except Exception as e:
        print(f"[ERROR] Gagal mengekspor file Excel: {e}")
        return False

def open_file_dialog() -> str:
    """Membuka dialog pemilih file Windows menggunakan Tkinter jika tersedia."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        file_path = filedialog.askopenfilename(
            title="Pilih File Excel / CSV Transaksi",
            filetypes=[("Excel & CSV files", "*.xlsx *.xls *.csv *.xlsm"), ("All files", "*.*")]
        )
        root.destroy()
        return file_path
    except Exception:
        return ""

def run_interactive_menu():
    """Menu interaktif ketika aplikasi dijalankan dengan double-click."""
    print("=" * 60)
    print("     APLIKASI REKAP & MAPPING TRANSAKSI EXCEL")
    print("=" * 60)
    print("1. Pilih & Rekap File Excel / CSV")
    print("2. Buat File Contoh (Sample Excel)")
    print("3. Keluar")
    print("=" * 60)
    
    pilihan = input("Pilih menu (1/2/3) [default: 1]: ").strip() or "1"
    
    if pilihan == "1":
        print("\nSilakan pilih file input:")
        print("- Ketik path file / drag & drop file ke jendela ini")
        print("- Atau tekan ENTER langsung untuk membuka jendela pemilih file (File Dialog)...")
        path_input = input("File: ").strip().strip('"').strip("'")
        
        if not path_input:
            path_input = open_file_dialog()
            if not path_input:
                print("[INFO] Tidak ada file yang dipilih.")
                return
        
        if not os.path.exists(path_input):
            print(f"[ERROR] File '{path_input}' tidak ditemukan.")
            return

        base, ext = os.path.splitext(path_input)
        default_out = f"{base}_rekap.xlsx"
        out_input = input(f"Nama file output [default: {os.path.basename(default_out)}]: ").strip().strip('"').strip("'")
        output_path = out_input if out_input else default_out

        tot_input = input("Tambahkan baris Total di akhir? (y/n) [default: y]: ").strip().lower()
        add_totals = (tot_input != "n")

        process_file(path_input, output_path, order="debet-first", add_totals=add_totals)

    elif pilihan == "2":
        sample_name = input("Nama file sample [default: sample_transaksi.xlsx]: ").strip() or "sample_transaksi.xlsx"
        count_str = input("Jumlah transaksi [default: 5]: ").strip() or "5"
        try:
            count = int(count_str)
        except ValueError:
            count = 5
        out_path = generate_sample_excel(sample_name, count=count)
        print(f"[SUKSES] File sample berhasil dibuat di: {out_path}")

    elif pilihan == "3":
        print("Sampai jumpa!")
        return

def main():
    # Cek jika file langsung di-drag-and-drop ke file .exe di Windows Explorer
    if len(sys.argv) == 2 and not sys.argv[1].startswith("-") and os.path.isfile(sys.argv[1]):
        input_file = sys.argv[1]
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_rekap.xlsx"
        print(f"[INFO] Mode Drag & Drop terdeteksi untuk: {input_file}")
        process_file(input_file, output_file, order="debet-first", add_totals=True)
        input("\nTekan ENTER untuk menutup jendela...")
        return

    # Jika tidak ada argumen sama sekali (double click exe)
    if len(sys.argv) == 1:
        try:
            run_interactive_menu()
        except KeyboardInterrupt:
            print("\n[INFO] Dibatalkan oleh pengguna.")
        input("\nTekan ENTER untuk keluar...")
        return

    # Mode CLI biasa dengan arguments
    parser = argparse.ArgumentParser(
        description="Script Python / EXE untuk merapikan & mapping transaksi Excel berdasarkan nomor BUKTI serta mengurutkan DEBET & KREDIT."
    )
    parser.add_argument("-i", "--input", type=str, help="Path file Excel/CSV input yang akan diproses")
    parser.add_argument("-o", "--output", type=str, default="hasil_rekap_transaksi.xlsx", help="Path file Excel output")
    parser.add_argument("--generate-sample", type=str, help="Buat file Excel contoh data acak")
    parser.add_argument("--count", type=int, default=5, help="Jumlah transaksi sampel (default: 5)")
    parser.add_argument("--order", choices=["debet-first", "kredit-first"], default="debet-first", help="Urutan DEBET/KREDIT")
    parser.add_argument("--add-totals", action="store_true", help="Sertakan baris total di akhir tabel")

    args = parser.parse_args()

    if args.generate_sample:
        print(f"[INFO] Membuat file sample Excel sebanyak {args.count} transaksi...")
        out_path = generate_sample_excel(args.generate_sample, count=args.count)
        print(f"[SUKSES] File sample berhasil dibuat di: {out_path}")
        return

    if not args.input:
        parser.print_help()
        sys.exit(1)

    success = process_file(args.input, args.output, order=args.order, add_totals=args.add_totals)
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m unittest discover tests`
Expected: All tests pass.

- [ ] **Step 5: Commit changes**

```bash
git add main.py tests/test_cli_interactive.py
git commit -m "feat: add interactive mode and drag-and-drop support to main"
```

---

### Task 2: Create PyInstaller Build Script and Specification

**Files:**
- Create: `build_exe.py`
- Create: `RekapTransaksi.spec`
- Test: `tests/test_build_config.py`

**Interfaces:**
- Consumes: PyInstaller CLI/API, `main.py`, `rekap_transaksi.py`
- Produces: `dist/RekapTransaksi.exe` standalone binary

- [ ] **Step 1: Write test for build script configuration**

```python
# tests/test_build_config.py
import unittest
import os

class TestBuildConfig(unittest.TestCase):
    def test_build_script_exists(self):
        self.assertTrue(os.path.exists("build_exe.py"))

    def test_spec_file_exists(self):
        self.assertTrue(os.path.exists("RekapTransaksi.spec"))

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests/test_build_config.py`
Expected: FAIL with `AssertionError: False is not true`

- [ ] **Step 3: Create `build_exe.py` and `RekapTransaksi.spec`**

Write `build_exe.py`:
```python
import os
import subprocess
import sys

def build():
    print("=" * 60)
    print("Memulai build standalone Windows EXE RekapTransaksi...")
    print("=" * 60)
    
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--name", "RekapTransaksi",
        "--collect-all", "openpyxl",
        "--collect-all", "pandas",
        "main.py"
    ]
    
    print(f"Menjalankan perintah: {' '.join(cmd)}")
    result = subprocess.run(cmd)
    
    if result.returncode == 0:
        exe_path = os.path.abspath(os.path.join("dist", "RekapTransaksi.exe"))
        print("\n" + "=" * 60)
        print(f"[SUKSES] File EXE berhasil dibuat di:\n{exe_path}")
        print("File ini siap disalin ke laptop lain tanpa perlu install Python.")
        print("=" * 60)
    else:
        print("\n[ERROR] Proses build gagal. Cek log di atas.")
        sys.exit(result.returncode)

if __name__ == "__main__":
    build()
```

- [ ] **Step 4: Run build script to generate executable and spec file**

Run: `python build_exe.py`
Expected: Successfully generates `dist/RekapTransaksi.exe` and `RekapTransaksi.spec`.

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m unittest tests/test_build_config.py`
Expected: PASS

- [ ] **Step 6: Commit build script & spec**

```bash
git add build_exe.py RekapTransaksi.spec tests/test_build_config.py
git commit -m "build: add pyinstaller packaging script and spec"
```

---

### Task 3: End-to-End Verification of Standalone EXE

**Files:**
- Test: `dist/RekapTransaksi.exe`
- Create: `dist/README_PANDUAN.txt` (User Guide for end users on other laptops)

**Interfaces:**
- Consumes: `dist/RekapTransaksi.exe`
- Produces: Validated standalone binary and instructions for end users

- [ ] **Step 1: Test running `--help` on the generated executable**

Run: `dist/RekapTransaksi.exe --help`
Expected: Output showing description and argument flags with exit code 0.

- [ ] **Step 2: Test running sample generation on the generated executable**

Run: `dist/RekapTransaksi.exe --generate-sample dist/sample_test.xlsx --count 5`
Expected: Output showing sample file created at `dist/sample_test.xlsx`.

- [ ] **Step 3: Test processing Excel file using the generated executable**

Run: `dist/RekapTransaksi.exe -i dist/sample_test.xlsx -o dist/hasil_rekap_test.xlsx --add-totals`
Expected: Output showing successful mapping and export to `dist/hasil_rekap_test.xlsx`.

- [ ] **Step 4: Create User Guide `dist/README_PANDUAN.txt`**

Write clear Indonesian instructions on how to use `RekapTransaksi.exe` on any Windows laptop (Double-click mode, Drag & drop mode, Command Prompt mode).

- [ ] **Step 5: Commit documentation and packaging changes**

```bash
git add dist/README_PANDUAN.txt
git commit -m "docs: add user manual for RekapTransaksi.exe"
```
