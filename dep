import pandas as pd
from os import path
import argparse

parser = argparse.ArgumentParser(description="Convert excel sheets to .dep files.")
parser.add_argument("-s", "--sheet", nargs="*", help="Name of sheets to be converted to .dep. Default will check for all sheets in excel file.")
parser.add_argument("-c", "--column", help="Set column number of non-decimal part of EMP_PAY", default=70, type=int)
parser.add_argument("-z", "--zeroes", help="Add inputted number of zeroes to the left of EMP_ACCTNO", default=0, type=int)
parser.add_argument("excel_file", help="Location of excel book")

args = parser.parse_args()

try:
    excel_book = pd.ExcelFile(args.excel_file)
except:
    print("Excel file not found or incorrect file location. Exiting...")
    exit(1)

if args.sheet is None:
    args.sheet = excel_book.sheet_names
else:
    for i in range(len(args.sheet)):
        try:
            args.sheet[i] = int(args.sheet[i])
        except:
            continue

for sheet in args.sheet:
    try:
        excel_df = pd.read_excel(excel_book, sheet_name=sheet, names=["A","B","C"], usecols="A:C", na_filter=False, header=None)
    except:
        print(f"Sheet \"{sheet}\" not found.")
        continue

    try:
        df_header_index = int(excel_df.index[excel_df["A"] == "EMP_ACCTNO"][0])
        for i in range(df_header_index+1):
            excel_df.drop(excel_df.index[0], inplace=True)
    except:
        if isinstance(sheet, int):
            print(f"EMP_ACCTNO not found in sheet no. {sheet}, \"{excel_book.sheet_names[sheet]}\"")
        else:
            print(f"EMP_ACCTNO not found in \"{sheet}\"")
        continue
    
    sheet_name = excel_book.sheet_names[sheet] if isinstance(sheet, int) else sheet
    with open(sheet_name + ".dep", "w", encoding="utf-8") as f:
        for i, row in excel_df.iterrows():
            try:
                f.write(f"{row['A']:>0{10+args.zeroes}}{row['B']:{50-args.zeroes+(args.column-70)}}{row['C']:{(10+2)}.2f}\n")
            except:
                continue
        print(f"Saved \"{sheet}\" to {path.abspath(f.name)}")