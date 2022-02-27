from typing import Optional, cast

import openpyxl as xl


class Materials:
    """excel list with layout Material:shelf in the warehouse"""

    def __init__(self, path: Optional[str] = None) -> None:
        if path == None:
            self.path = "../excel_files/stock.xlsx"
        else:
            self.path = cast(str, path)
        self.wb = xl.load_workbook(self.path, data_only=True)
        self.materials_locations: dict[str, int] = self.get_material_list()

    def get_material_list(self) -> dict:
        ws = self.wb.active
        products: dict = {}
        for column in ws.iter_cols(min_col=1, max_col=1):
            for cell in column:
                products[cell.value] = cell.offset(0, 1).value

        return products
