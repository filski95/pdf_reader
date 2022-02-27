import os
from io import StringIO
from typing import Optional

from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser


def get_pdf_name(path: Optional[str] = None):
    if path == None:
        path = "../excel_files"

    found_files = [file for file in os.listdir(path) if file.endswith(".pdf")]
    return found_files


def read_pdf(file_name: str, path=None) -> list[str]:
    # *https://pdfminersix.readthedocs.io/en/latest/tutorial/composable.html
    if path is None:
        path = os.path.join("../excel_files", file_name)
    output_string = StringIO()
    with open(path, "rb") as f:
        parser = PDFParser(f)
        doc = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)
    string = output_string.getvalue()
    list_of_strings = string.split("\n")
    return list_of_strings


def get_locations(material_list: dict, picklist: list) -> dict:
    dict_materials_picklist: dict = {}

    for material in picklist:
        if material in material_list.keys():
            dict_materials_picklist[material] = material_list[material]
    return dict_materials_picklist
