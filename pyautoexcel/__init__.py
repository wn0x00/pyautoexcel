import win32com.client


class Excel(object):
    def __init__(self, filename=None):
        self.application = win32com.client.Dispatch("Excel.Application")
        self.application.Visible = 1
        if filename:
            self.filename = filename
            self.workbook = self.application.Workbooks.Open(filename)
        else:
            self.filename = ""
            self.workbook = self.application.Workbooks.Add()


    def save(self, new_filename):
        if new_filename:
            self.filename = new_filename
            self.workbook.SaveAs(new_filename)
        else:
            self.workbook.Save()

    def close(self):
        self.workbook.Close(SaveChanges=0)
        del self.application
