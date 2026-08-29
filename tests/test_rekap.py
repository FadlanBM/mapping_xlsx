import os
import unittest
import pandas as pd
from rekap_transaksi import clean_and_standardize, sort_by_bukti_and_type, generate_sample_excel, export_to_excel

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

    def test_generate_sample_and_export(self):
        sample_path = "tests/test_sample.xlsx"
        out_path = "tests/test_output.xlsx"
        try:
            generate_sample_excel(sample_path, count=3)
            self.assertTrue(os.path.exists(sample_path))
            
            df = pd.read_excel(sample_path)
            sorted_df = sort_by_bukti_and_type(df)
            export_to_excel(sorted_df, out_path)
            self.assertTrue(os.path.exists(out_path))
        finally:
            if os.path.exists(sample_path):
                os.remove(sample_path)
            if os.path.exists(out_path):
                os.remove(out_path)

    def test_export_auto_adds_xlsx_extension(self):
        df = pd.DataFrame({
            "BUKTI": ["B01"],
            "NAMA AKUN": ["Kas"],
            "DEBET": [1000],
            "KREDIT": [0],
            "KETERANGAN": ["Test"]
        })
        # Test export without extension
        out_no_ext = "tests/test_no_ext_export"
        expected_file = "tests/test_no_ext_export.xlsx"
        try:
            res = export_to_excel(df, out_no_ext)
            self.assertTrue(res.endswith(".xlsx"))
            self.assertTrue(os.path.exists(expected_file))
        finally:
            if os.path.exists(expected_file):
                os.remove(expected_file)

        # Test generate sample without extension
        sample_no_ext = "tests/test_sample_no_ext"
        expected_sample = "tests/test_sample_no_ext.xlsx"
        try:
            res_sample = generate_sample_excel(sample_no_ext, count=2)
            self.assertTrue(res_sample.endswith(".xlsx"))
            self.assertTrue(os.path.exists(expected_sample))
        finally:
            if os.path.exists(expected_sample):
                os.remove(expected_sample)

if __name__ == "__main__":
    unittest.main()
