# Rekap Transaksi Excel Berdasarkan No BUKTI & Urutan Debet/Kredit Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Membaca data transaksi dengan kolom `NAMA AKUN`, `BUKTI`, `DEBET`, `KREDIT`, `KETERANGAN`, mengurutkan data berdasarkan nomor `BUKTI`, menyusun baris DEBET dan KREDIT beserta keterangan secara berurutan pada BUKTI yang sama, dan mengekspor hasilnya kembali ke file Excel dengan styling profesional.

**Architecture:** 
- `rekap_transaksi.py`: Logika standardisasi kolom, sorting berdasarkan BUKTI dan Debet/Kredit, generator sampel data, dan exporter styling openpyxl.
- `main.py`: CLI script untuk user.
- `tests/test_rekap.py`: Automated testing suite.

**Tech Stack:** Python 3.8+, pandas, openpyxl.

## Global Constraints
- Kolom output wajib: `NAMA AKUN`, `BUKTI`, `DEBET`, `KREDIT`, `KETERANGAN`.
- Penanganan angka pada `DEBET` dan `KREDIT` yang fleksibel (bisa menangani format string dengan koma/titik, NaN, atau angka desimal).
- Pada `BUKTI` yang sama, baris dengan DEBET > 0 diposisikan di atas baris dengan KREDIT > 0.
- File Excel hasil ekspor terformat rapi (Header styling, format angka mata uang `Rp #,##0`, border, auto width).

---

### Task 1: Setup Core Transaction Logic and Column Mapping

**Files:**
- Create: `requirements.txt`
- Create: `rekap_transaksi.py`
- Create: `tests/test_rekap.py`

**Interfaces:**
- Produces:
  - `clean_and_standardize(df: pd.DataFrame) -> pd.DataFrame`
  - `sort_by_bukti_and_type(df: pd.DataFrame, debet_first: bool = True) -> pd.DataFrame`

- [ ] **Step 1: Write the failing tests in `tests/test_rekap.py`**

```python
import unittest
import pandas as pd
from rekap_transaksi import clean_and_standardize, sort_by_bukti_and_type

class TestRekapTransaksi(unittest.TestCase):
    def test_clean_and_standardize(self):
        df = pd.DataFrame({
            "Nama Akun": ["Kas", "Persediaan"],
            "bukti": ["BKT-001", "BKT-001"],
            "debet": ["100.000", 0],
            "kredit": [0, "100.000"],
            "keterangan": ["Kas Masuk", "Beli Barang"]
        })
        res = clean_and_standardize(df)
        expected_cols = ["NAMA AKUN", "BUKTI", "DEBET", "KREDIT", "KETERANGAN"]
        self.assertEqual(list(res.columns), expected_cols)
        self.assertEqual(res["DEBET"].iloc[0], 100000.0)

    def test_sort_by_bukti_and_type(self):
        # Input with unordered BUKTI and Kredit placed before Debet
        df = pd.DataFrame({
            "NAMA AKUN": ["Utang Usaha", "Persediaan Barang", "Kas", "Beban Listrik"],
            "BUKTI": ["BKT-002", "BKT-002", "BKT-001", "BKT-001"],
            "DEBET": [0, 500000, 0, 150000],
            "KREDIT": [500000, 0, 150000, 0],
            "KETERANGAN": ["Pelunasan Utang", "Pembelian Barang", "Bayar Listrik", "Listrik Kantor"]
        })
        sorted_df = sort_by_bukti_and_type(df, debet_first=True)
        
        # Verify BUKTI ordering (BKT-001 then BKT-002)
        # And within BKT-001: Debet (Beban Listrik) then Kredit (Kas)
        self.assertEqual(sorted_df["BUKTI"].iloc[0], "BKT-001")
        self.assertEqual(sorted_df["NAMA AKUN"].iloc[0], "Beban Listrik") # Debet
        self.assertEqual(sorted_df["DEBET"].iloc[0], 150000)
        
        self.assertEqual(sorted_df["BUKTI"].iloc[1], "BKT-001")
        self.assertEqual(sorted_df["NAMA AKUN"].iloc[1], "Kas") # Kredit
        self.assertEqual(sorted_df["KREDIT"].iloc[1], 150000)

        # Within BKT-002: Debet (Persediaan Barang) then Kredit (Utang Usaha)
        self.assertEqual(sorted_df["BUKTI"].iloc[2], "BKT-002")
        self.assertEqual(sorted_df["NAMA AKUN"].iloc[2], "Persediaan Barang") # Debet
        self.assertEqual(sorted_df["BUKTI"].iloc[3], "BKT-002")
        self.assertEqual(sorted_df["NAMA AKUN"].iloc[3], "Utang Usaha") # Kredit

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests/test_rekap.py`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Implement `requirements.txt` and `rekap_transaksi.py`**

Create `requirements.txt`:
```txt
pandas>=2.0.0
openpyxl>=3.1.0
```

Implement in `rekap_transaksi.py`:
```python
import pandas as pd
import numpy as np
import re

def _clean_numeric(val):
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace("Rp", "").replace("rp", "").strip()
    if not val_str or val_str == "-":
        return 0.0
    # Clean thousand separators
    if "." in val_str and "," in val_str:
        # e.g., 1.000,50 -> 1000.50
        val_str = val_str.replace(".", "").replace(",", ".")
    elif "." in val_str:
        # Could be 1.000 or 10.5
        parts = val_str.split(".")
        if len(parts[-1]) == 3: # likely thousands
            val_str = val_str.replace(".", "")
    elif "," in val_str:
        val_str = val_str.replace(",", "")
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def clean_and_standardize(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    col_mapping = {}
    for col in df.columns:
        clean_col = re.sub(r'[^a-zA-Z0-9]', '', str(col).strip().upper())
        if 'AKUN' in clean_col or 'ACCOUNT' in clean_col or 'REKENING' in clean_col:
            col_mapping[col] = 'NAMA AKUN'
        elif 'BUKTI' in clean_col or 'NOBUKTI' in clean_col or 'VOUCHER' in clean_col or 'INVOICE' in clean_col or 'TRX' in clean_col or 'ID' in clean_col:
            col_mapping[col] = 'BUKTI'
        elif 'DEBET' in clean_col or 'DEBIT' in clean_col:
            col_mapping[col] = 'DEBET'
        elif 'KREDIT' in clean_col or 'CREDIT' in clean_col:
            col_mapping[col] = 'KREDIT'
        elif 'KETERANGAN' in clean_col or 'KET' in clean_col or 'DESC' in clean_col or 'DESKRIPSI' in clean_col or 'CATATAN' in clean_col:
            col_mapping[col] = 'KETERANGAN'
            
    df.rename(columns=col_mapping, inplace=True)
    
    target_cols = ['NAMA AKUN', 'BUKTI', 'DEBET', 'KREDIT', 'KETERANGAN']
    for col in target_cols:
        if col not in df.columns:
            df[col] = "" if col not in ['DEBET', 'KREDIT'] else 0.0
            
    # Clean numerical columns
    df['DEBET'] = df['DEBET'].apply(_clean_numeric)
    df['KREDIT'] = df['KREDIT'].apply(_clean_numeric)
    df['BUKTI'] = df['BUKTI'].astype(str).str.strip()
    df['NAMA AKUN'] = df['NAMA AKUN'].astype(str).str.strip()
    df['KETERANGAN'] = df['KETERANGAN'].astype(str).str.strip()
    
    return df[target_cols]

def sort_by_bukti_and_type(df: pd.DataFrame, debet_first: bool = True) -> pd.DataFrame:
    df = clean_and_standardize(df)
    
    # Priority: Debet row (DEBET > 0) -> 1, Kredit row (KREDIT > 0) -> 2
    # If debet_first is False, reverse the priority
    def get_order_priority(row):
        d_val = row['DEBET']
        k_val = row['KREDIT']
        if d_val > 0 and k_val == 0:
            return 1 if debet_first else 2
        elif k_val > 0 and d_val == 0:
            return 2 if debet_first else 1
        elif d_val > 0 and k_val > 0:
            return 1.5 # both
        return 3

    df['_sort_priority'] = df.apply(get_order_priority, axis=1)
    
    # Sort by BUKTI then sort_priority
    df_sorted = df.sort_values(by=['BUKTI', '_sort_priority'], ascending=[True, True]).copy()
    df_sorted.drop(columns=['_sort_priority'], inplace=True)
    return df_sorted.reset_index(drop=True)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m unittest tests/test_rekap.py`
Expected: PASS

---

### Task 2: Excel Export Styling and Sample Generator

**Files:**
- Modify: `rekap_transaksi.py`
- Modify: `tests/test_rekap.py`

**Interfaces:**
- Produces:
  - `export_to_excel(df: pd.DataFrame, output_path: str) -> str`
  - `generate_sample_excel(output_path: str, count: int = 10) -> str`

- [ ] **Step 1: Write tests for Excel exporter and sample generator**
- [ ] **Step 2: Implement styling in `openpyxl` with currency format and borders**
- [ ] **Step 3: Run tests to verify pass**

---

### Task 3: CLI Runner Script

**Files:**
- Create: `main.py`
- Create: `README.md`

- [ ] **Step 1: Create `main.py` CLI supporting `--input`, `--output`, and `--generate-sample`**
- [ ] **Step 2: Verify end-to-end processing with sample generation and re-sorting**
