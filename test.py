import openpyxl as xl
from openpyxlpractice import check_data
wb = xl.load_workbook("book3.")
ws = wb.active
check_data(ws)