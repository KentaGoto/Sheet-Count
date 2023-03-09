import openpyxl
import sys


def check_argument(arg_l, f):
    if arg_l == 0:
        print("Script file name.")
    elif arg_l == 1:
        print("Argument is 1.")
        return input("File: ")
    elif arg_l == 2:
        print("Argument is 2.")
        f = sys.argv[1]
        return f
    else:
        print("Error: The argument can be neither 1 nor 2. Forced termination.")
        sys.exit()


if __name__ == '__main__':
    arg_l = len(sys.argv)
    f = ''
    f = check_argument(arg_l, f)

    # Loading Workbooks
    workbook = openpyxl.load_workbook(f)

    # Retrieve and output sheet name and number of sheets.
    sheet_names = workbook.sheetnames
    sheet_count = len(sheet_names)
    print("Number of sheets: ", sheet_count)
    for i, sheet_name in enumerate(sheet_names, start=1):
        print("Sheet", i, ": ", sheet_name)
