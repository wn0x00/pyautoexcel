from typing import Any
import win32com.client


class Excel(object):
    def __init__(self, filename=None) -> None:
        self.application = win32com.client.Dispatch("Excel.Application")
        self.application.Visible = 1
        if filename:
            self.filename = filename
            self.workbook = self.application.Workbooks.Open(filename)
        else:
            self.filename = ""
            self.workbook = self.application.Workbooks.Add()

    def save(self, new_filename=None) -> None:
        if new_filename:
            self.filename = new_filename
            self.workbook.SaveAs(new_filename)
        else:
            self.workbook.Save()

    def close(self) -> None:
        self.workbook.Close(SaveChanges=0)

    def get_cell(self, sheet: str, row: int, col: str) -> Any:
        sht = self.application.Worksheets(sheet)
        return sht.Cells(row, col).Value

    def set_cell(self, sheet: str, row: int, col: str, value: Any) -> None:
        sht = self.application.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def get_range(
        self,
        sheet: str,
        row_top: int,
        col_left: str,
        row_bot: int,
        col_right: str,
    ) -> list:
        sht = self.application.Worksheets(sheet)
        return sht.Range(
            sht.Cells(row_top, col_left), sht.Cells(row_bot, col_right)
        ).Value

    def set_range(self, sheet: str, top_row: int, left_col: str, data: list) -> None:
        bottom_row = top_row + len(data) - 1
        right_col = top_row + len(data[0]) - 1
        sht = self.application.Worksheets(sheet)
        sht.Range(
            sht.Cells(top_row, left_col), sht.Cells(bottom_row, right_col)
        ).Value = data

    def get_sheetnames(self) -> list:
        count = self.application.Sheets.Count
        return [self.application.Sheets[i].Name for i in range(count)]

    def add_sheet(self, name: str, before=True) -> None:
        if before:
            before, after = True, False
        else:
            before, after = True, False
        self.application.Sheets.Add(Before=before, After=after, Count=1, Type=-4167)
        self.workbook.ActiveSheet.Name = name

    def delete_sheet(self) -> None:
        pass


if __name__ == "__main__":
    excel = Excel()
    excel.add_sheet("new_sheet")
