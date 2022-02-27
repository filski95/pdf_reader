from cgitb import small
from black import put_trailing_semicolon_back
from macro_files.materials import Materials
from macro_files import pdf_reader
from macro_files import output_file


def run_macro():
    materials = Materials()
    small_components_list_pdf = pdf_reader.get_pdf_name()
    for list in small_components_list_pdf:

        picklist = pdf_reader.read_pdf(list)
        picklist_locations = pdf_reader.get_locations(materials.materials_locations, picklist)
        output = output_file.OutputFile([picklist_locations], list)
        output.dump_data()


if __name__ == "__main__":
    run_macro()
