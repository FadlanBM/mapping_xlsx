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
        print("\nSilakan tentukan file input:")
        print("- Ketik / drag & drop file ke jendela ini")
        print("- Atau tekan ENTER langsung untuk membuka jendela pemilih file...")
        path_input = input("Path file: ").strip().strip('"').strip("'")
        
        if not path_input:
            path_input = open_file_dialog()
            if not path_input:
                print("[INFO] Tidak ada file yang dipilih.")
                return
            print(f"File dipilih: {path_input}")
        
        if not os.path.exists(path_input):
            print(f"[ERROR] File '{path_input}' tidak ditemukan.")
            return

        base, ext = os.path.splitext(path_input)
        default_out = f"{base}_rekap.xlsx"
        out_input = input(f"Nama file output [default: {os.path.basename(default_out)}]: ").strip().strip('"').strip("'")
        if out_input:
            if not out_input.lower().endswith('.xlsx'):
                out_input = f"{os.path.splitext(out_input)[0]}.xlsx"
            if not os.path.isabs(out_input) and not os.path.dirname(out_input):
                input_dir = os.path.dirname(os.path.abspath(path_input))
                output_path = os.path.join(input_dir, out_input)
            else:
                output_path = out_input
        else:
            output_path = default_out

        tot_input = input("Tambahkan baris Total di akhir? (y/n) [default: y]: ").strip().lower()
        add_totals = (tot_input != "n")

        process_file(path_input, output_path, order="debet-first", add_totals=add_totals)

    elif pilihan == "2":
        sample_name = input("Nama file sample [default: sample_transaksi.xlsx]: ").strip().strip('"').strip("'") or "sample_transaksi.xlsx"
        if not sample_name.lower().endswith('.xlsx'):
            sample_name = f"{os.path.splitext(sample_name)[0]}.xlsx"
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
