import openpyxl

def get_active_sheet():
    filename = "TrackerFile.xlsx"
    wb = openpyxl.load_workbook(filename)

    # Retrieve names of all worksheet
    sheet_names = wb.sheetnames

    # Catch index location of sheet
    Associated_Sheet_number_str = ""

    for i,sheet in enumerate(sheet_names):
        Associated_Sheet_number_str += f"\n {i} ---> {sheet}"

    access_sheet = ""
    # Select active worksheet
    access_sheet = input(f"\n Enter the associated number for the following sheet {Associated_Sheet_number_str}")

    # Test for valid input
    if access_sheet.isnumeric() != True:
        print("Key provided is not a string. Please try again!")
        access_sheet = "FLAG_Negative"
        pass
    else:
        access_sheet = int(access_sheet)

    if access_sheet == "FLAG_Negative":
        print("error_No_active_Sheet")
        active_sheet = "String"
    else:
        Index_of_Sheet_of_interest = access_sheet
        active_sheet = wb[sheet_names[Index_of_Sheet_of_interest]]

    return filename, wb, active_sheet, Index_of_Sheet_of_interest