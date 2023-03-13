import openpyxl
import sys
import warnings
import os


def check_argument(arg_l):
    if arg_l == 0:
        print("Script file name.")
    elif arg_l == 1:
        return input("File: ")
    elif arg_l == 2:
        return sys.argv[1]
    else:
        print("Error: Invalid number of arguments. Forced termination.")
        sys.exit()


if __name__ == '__main__':
    arg_l = len(sys.argv)
    f = check_argument(arg_l)

    # Loading Workbooks
    warnings.simplefilter("ignore")
    workbook = openpyxl.load_workbook(f)

    # Retrieve and output sheet name and number of sheets.
    sheet_names = workbook.sheetnames
    sheet_count = len(sheet_names)
    print(f)
    print("")
    print("Number of sheets: ", sheet_count)
    for i, sheet_name in enumerate(sheet_names, start=1):
        print("Sheet", i, ": ", sheet_name)
    
    os.system("pause")
