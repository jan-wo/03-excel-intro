from src.excel_file import ExcelFile
import src.utils as utl


CONF = utl.get_yaml('config.yml')

test_file = ExcelFile()
test_file.add_workbook(CONF['input_path'])
test_file.select_sheet(CONF['sheet_name'])
test_file.data_from_range(CONF['get_data_cell_address'])
test_file.data_to_range(CONF['data_to_insert'], CONF['set_data_cell_address'])
test_file.save_close_wb(CONF['output_path'])
test_file.stop_app()
