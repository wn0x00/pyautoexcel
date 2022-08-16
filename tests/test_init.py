import os
import unittest

from pyautoexcel import Excel


class ExcelBase(unittest.TestCase):
    def test_open_file(self):
        data_path = os.path.dirname(__file__)
        file_path = os.path.join(data_path, "data", "demo.xlsx")
        excel_ref = Excel()
        result = excel_ref.open(file_path)
        self.assertIsNone(result)
