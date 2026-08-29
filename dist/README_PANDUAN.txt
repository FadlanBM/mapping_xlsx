========================================================================
             PANDUAN PENGGUNAAN APLIKASI REKAP TRANSAKSI (EXE)
========================================================================

Aplikasi ini sudah dikemas menjadi standalone file EXE (RekapTransaksi.exe),
sehingga DAPAT DIJALANKAN LANGSUNG di laptop/komputer Windows apapun
TANPA PERLU menginstall Python maupun aplikasi tambahan lainnya.

------------------------------------------------------------------------
CARA PENGGUNAAN:
------------------------------------------------------------------------

CARA 1: DOUBLE CLICK (Paling Mudah)
1. Double click file `RekapTransaksi.exe`.
2. Jendela menu interaktif akan muncul.
3. Pilih opsi [1] lalu tekan ENTER untuk membuka dialog pemilih file (File Explorer).
4. Pilih file Excel (.xlsx, .xls) atau .csv yang ingin direkap.
5. File hasil rekap akan otomatis disimpan di folder yang sama dengan nama tambahan `_rekap.xlsx`.

CARA 2: DRAG & DROP (Tarik dan Lepas)
1. Cukup drag (tarik) file Excel Anda dari Windows Explorer dan drop (lepas)
   langsung ke atas ikon file `RekapTransaksi.exe`.
2. Aplikasi akan otomatis memproses dan menyimpan hasil rekapnya.

CARA 3: MENGGUNAKAN COMMAND PROMPT / POWERSHELL
Buka Command Prompt (CMD) di folder ini, lalu jalankan perintah:

- Memproses file:
    RekapTransaksi.exe -i "nama_file_input.xlsx" -o "hasil_rekap.xlsx" --add-totals

- Membuat contoh file transaksi:
    RekapTransaksi.exe --generate-sample "contoh_data.xlsx" --count 10

- Menampilkan bantuan perintah:
    RekapTransaksi.exe --help

========================================================================
