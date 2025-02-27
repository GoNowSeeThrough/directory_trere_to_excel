import unittest
import os
import pandas as pd
from handover_list import generate_handover_list
import handover_list

class TestHandoverList(unittest.TestCase):

    def setUp(self):
        # Set up a test directory and files
        self.test_dir = os.path.join(os.getcwd(), "test_data") # Create test_data within current directory
        if os.path.exists(self.test_dir):
            import shutil
            shutil.rmtree(self.test_dir)
        os.makedirs(self.test_dir, exist_ok=True)
        os.makedirs(os.path.join(self.test_dir, "subfolder1"), exist_ok=True)
        os.makedirs(os.path.join(self.test_dir, "subfolder2"), exist_ok=True)
        with open(os.path.join(self.test_dir, "file1.txt"), "w") as f:
            f.write("test content")
        with open(os.path.join(self.test_dir, "subfolder1", "file2.txt"), "w") as f:
            f.write("test content")
        self.excel_file = os.path.join(self.test_dir, "交接清单2025-02-27.xlsx") # Fixed date for testing

    def tearDown(self):
        # Clean up the test directory and files
        if os.path.exists(self.excel_file):
            os.remove(self.excel_file)
        import shutil
        try:
            shutil.rmtree(self.test_dir)
        except FileNotFoundError:
            pass # Handle case where test_dir is already deleted

    def test_excel_file_creation(self):
        # Test that the Excel file is created
        os.chdir(self.test_dir)
        generate_handover_list()
        self.assertTrue(os.path.exists(self.excel_file))
        os.chdir("..")

    def test_excel_data(self):
        # Test that the Excel file contains the correct data
        os.chdir(self.test_dir)
        generate_handover_list()
        df = pd.read_excel(self.excel_file)
        self.assertEqual(len(df), 3) # 3 files + 1 empty subfolder2
        self.assertEqual(df.iloc[0]["项目名称"], "test_data")
        self.assertTrue(pd.isna(df.iloc[0]["内容模块"]))
        self.assertEqual(df.iloc[0]["文件名"], "file1.txt")
        self.assertEqual(df.iloc[1]["内容模块"], "subfolder1")
        self.assertEqual(df.iloc[1]["文件名"], "file2.txt")
        self.assertEqual(df.iloc[2]["内容模块"], "subfolder2")
        self.assertTrue(pd.isna(df.iloc[2]["文件名"])) #empty subfolder
        os.chdir("..")

    def test_empty_directory(self):
        # Test that the script handles an empty directory correctly
        os.chdir(".")
        os.makedirs("empty_test_dir", exist_ok=True)
        os.chdir("empty_test_dir")
        generate_handover_list()
        self.assertTrue(os.path.exists(os.path.join(os.getcwd(), "交接清单2025-02-27.xlsx")))
        df = pd.read_excel(os.path.join(os.getcwd(), "交接清单2025-02-27.xlsx"))
        self.assertEqual(len(df), 0)
        os.chdir("..")
        import shutil
        shutil.rmtree("empty_test_dir")

if __name__ == '__main__':
    unittest.main()
