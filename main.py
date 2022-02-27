from collections import defaultdict


from macro_files import output_file, pdf_reader
from macro_files.excel_formating import create_combined_excel
from macro_files.materials import Materials


def run_macro() -> None:
    materials = Materials()
    small_components_list_pdf = pdf_reader.get_pdf_name()
    for list in small_components_list_pdf:

        picklist = pdf_reader.read_pdf(list)
        picklist_locations = pdf_reader.get_locations(materials.materials_locations, picklist)
        output = output_file.OutputFile([picklist_locations], list)
        output.dump_data()

    create_combined_excel()


if __name__ == "__main__":
    run_macro()
