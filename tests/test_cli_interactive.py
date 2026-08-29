import unittest
import os
import sys
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
            try:
                os.remove(self.sample_input)
            except OSError:
                pass
        if os.path.exists(self.sample_output):
            try:
                os.remove(self.sample_output)
            except OSError:
                pass

    def test_process_file_success(self):
        success = process_file(self.sample_input, self.sample_output, order="debet-first", add_totals=True)
        self.assertTrue(success)
        self.assertTrue(os.path.exists(self.sample_output))

if __name__ == "__main__":
    unittest.main()
