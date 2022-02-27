import openpyxl as xl
import os
from openpyxl.styles import Font


def create_combined_excel():
    # * comining all single files into one, to ease printing of labels
    dir = "../excel_files/output_files/"
    files = os.listdir(dir)
    wb = xl.Workbook()
    ws = wb.active
    combine_excels(dir, ws, files)
    highlight_za(ws)
    wb.save("../excel_files/output_files/combined.xlsx")


def combine_excels(dir, ws, files) -> None:
    for file in files:
        if file.endswith(".xlsx"):
            file_wb = xl.load_workbook(filename=os.path.join(dir + file))
            file_ws = file_wb.active
            for value in file_ws.values:  # * generator
                ws.append(value)


def highlight_za(ws):
    # * styling of cells with ZAs.
    za_cell_style = Font(bold=True, color="FF0000", size=14)
    for row in ws.iter_rows(min_row=1, max_col=1):
        for cell in row:
            if str(cell.value).startswith("ZA"):
                cell.font = za_cell_style
