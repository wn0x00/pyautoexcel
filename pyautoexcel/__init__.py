import os
import win32com.client


class Excel(object):
    def __init__(self):
        self._excel = win32com.client.Dispatch("Excel.Application")
        self._excel.Visible = 1

    def open(self, filename: str) -> None:
        self._excel.Workbooks.Open(filename)
