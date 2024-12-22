import os
import openpyxl as xl
import datetime as dt



EXCEL_FILE = 'attendance_sheet.xlsx'


def get_session():
    current_hour = dt.datetime.now().hour
    if current_hour < 12:
        return 'M'
    else:
        return 'E'


def add_attendance(roll_number):
    wb = xl.load_workbook(EXCEL_FILE)
    sheet = wb.active

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=1):
        if row[0].value == roll_number:
            if get_session() == 'M':
                sheet[f'B{row[0].row}'] = 'Present'
                sheet[f'C{row[0].row}'] = 'Present'
            else:
                sheet[f'C{row[0].row}'] = 'Present'

    wb.save(EXCEL_FILE)
    wb.close()
    return f"Attendance marked for roll number {roll_number}"



if __name__ == '__main__':
    print(add_attendance('23CE1017'))