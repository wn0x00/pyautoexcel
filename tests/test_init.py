import os
import unittest
import time

from pyautoexcel import Excel


class ExcelBase(unittest.TestCase):
    def test_init_file(self):
        data_path = os.path.dirname(__file__)
        file_path = os.path.join(data_path, "data", "demo.xlsx")
        excel_ref = Excel()
        time.sleep(3)
        self.assertIsNotNone(excel_ref)
        excel_ref.close()
        excel_ref.application.Quit()

        excel_ref = Excel(file_path)
        self.assertIsNotNone(excel_ref)
        time.sleep(3)
        excel_ref.close()
        excel_ref.application.Quit()

    def test_close(self):
        excel_ref = Excel()
        time.sleep(3)
        self.assertIsNotNone(excel_ref)
        excel_ref.close()
        excel_ref.application.Quit()

    def test_save(self):
        data_path = os.path.dirname(__file__)
        file_path = os.path.join(data_path, "data", "demo.xlsx")
        excel_ref = Excel(file_path)
        time.sleep(3)
        self.assertIsNotNone(excel_ref)
        excel_ref.save()
        file_path = os.path.join(data_path, "data", "demo2.xlsx")
        excel_ref.save(file_path)
        excel_ref.close()
        excel_ref.application.Quit()

    def test_set_cell(self):
        excel_ref = Excel()
        excel_ref.set_cell("Sheet1", 1, "A", "1111")
        excel_ref.close()
        excel_ref.application.Quit()

    def test_get_cell(self):
        excel_ref = Excel()
        excel_ref.set_cell("Sheet1", 1, "A", "1111")
        res = excel_ref.get_cell("Sheet1", 1, "A")
        self.assertEqual(res, 1111)
        excel_ref.close()
        excel_ref.application.Quit()

    def test_set_range(self):
        excel_ref = Excel()
        excel_ref.set_range("Sheet1", 1, "A", [[1, 2], [3, 4]])
        excel_ref.close()
        excel_ref.application.Quit()

    def test_get_range(self):
        excel_ref = Excel()
        excel_ref.set_range("Sheet1", 1, "A", [[1, 2], [3, 4]])
        res = excel_ref.get_range("Sheet1", 1, "A", 2, "B")
        self.assertEqual(res, ((1.0, 2.0), (3.0, 4.0)))
        excel_ref.close()
        excel_ref.application.Quit()
