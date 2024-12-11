import sys
import openpyxl

def get_excel_sheet_names(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet_names = workbook.sheetnames
        for sheet_name in sheet_names:
            print(sheet_name)
        workbook.close()
    except Exception as e:
        print("Error:", e)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python get_excel_sheet_names.py <file_path>")
        sys.exit(1)
    
    file_path = sys.argv[1]
    get_excel_sheet_names(file_path)
