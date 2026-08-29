import os
import subprocess
import sys

def build():
    print("=" * 60)
    print("Memulai build standalone Windows EXE: RekapTransaksi...")
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
    
    print(f"Menjalankan perintah:\n{' '.join(cmd)}\n")
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
