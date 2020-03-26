import xlwings as xw


class ExcelFile():

    def __init__(self, visible=False):
        print("[ExcelFile]: object created")
        self._app = xw.App(visible=visible)

    def stop_app(self):
        print(f'[ExcelFile]: Closing application.')
        self._app.quit()

    def add_workbook(self, wb_path):
        print(f'[ExcelFile]: Adding workbook: <{wb_path}>')
        self._wb = self._app.books.open(wb_path)

        print(f'[ExcelFile]: Workbooks assigned to app: {self._app.books}')

    def save_close_wb(self, new_wb_path):
        print(f'[ExcelFile]: Saving application as <{new_wb_path}>')
        self._wb.save(new_wb_path)
        self._wb.close()

    def select_sheet(self, sheet_name):
        self._sh = self._wb.sheets(sheet_name)
        print(f'[ExcelFile]: Selected sheet <{self._sh}>')

    def data_from_range(self, address):
        data = self._sh.range(address).value
        print(f'[ExcelFile]: Data taken from <{address}>: <{data}>')

    def data_to_range(self, data, address):
        self._sh.range(address).value = data
        print(f'[ExcelFile]: Inserting <{data}> to <{address}>')
