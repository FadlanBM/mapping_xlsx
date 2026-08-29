import unittest
import os

class TestBuildConfig(unittest.TestCase):
    def test_build_script_exists(self):
        self.assertTrue(os.path.exists("build_exe.py"))

    def test_spec_file_exists(self):
        self.assertTrue(os.path.exists("RekapTransaksi.spec"))

if __name__ == "__main__":
    unittest.main()
