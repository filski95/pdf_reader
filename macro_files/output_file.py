import os
from typing import Optional

import openpyxl as xl
import pandas as pd
from genericpath import exists
from openpyxl import Workbook


class OutputFile:
    def __init__(self, data: list, file_name: str) -> None:
        self.file_name = file_name
        self.new_file_path = os.path.join("../excel_files/output_files", self.file_name[:-4] + ".xlsx")
        self.wb: Workbook = xl.Workbook()
        self.data = data
        self.format_file()
        self.wb.save(filename=self.new_file_path)

    def dump_data(self) -> None:
        df = pd.DataFrame(data=self.data)
        df = df.T
        # * to append ExcelWriter needs to be used.
        with pd.ExcelWriter(self.new_file_path, mode="a", if_sheet_exists="overlay") as writer:
            df.to_excel(writer, sheet_name="sheet1", startrow=1, index=True, header=False)

    def format_file(self) -> None:
        ws = self.wb.active
        # * using sheet1, the same as in dump_data to put data into the same sheet + overlay option
        ws.title = "sheet1"
        ws["A1"] = str(self.file_name[:-4])
        ws.sheet_format.baseColWidth = 15  # * make columns wider
