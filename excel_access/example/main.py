import os
import sys
parent_dir_name = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
sys.path.append(parent_dir_name)
from excel_access import ExcelAccessor


def main():
    excel_handle = ExcelAccessor("./test.xlsx")
    excel_handle.write_cell("test123", 1, "A", "1A")
    excel_handle.write_cell("test123", 1, "B", "1B")
    print("Max row number of sheet test123 is " + str(excel_handle.get_max_row("test123")) + "\n")
    print("Max column number of sheet test123 is " + str(excel_handle.get_max_column("test123")) + "\n")
    print("cell 1A is " + str(excel_handle.read_cell("test123", 1, "A")) + "\n")
    print("cell 1B is " + str(excel_handle.read_cell("test123", 1, "B")) + "\n")


if __name__ == '__main__':
    main()
