import os
import subprocess
import sys

def build():
    print("=" * 60)
    print("Memulai build standalone Windows EXE: RekapTransaksi...")
    print("=" * 60)

    # Tutup instance RekapTransaksi.exe yang mungkin masih berjalan
    if sys.platform == "win32":
        subprocess.run(["taskkill", "/F", "/IM", "RekapTransaksi.exe", "/T"], 
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    dist_exe = os.path.join("dist", "RekapTransaksi.exe")
    if os.path.exists(dist_exe):
        try:
            os.remove(dist_exe)
        except OSError:
            pass
    
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
