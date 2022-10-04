import os
from openpyxl import load_workbook

class print_pls:

    def make_file(date1, date2):
        print_file = open(f"{date1.date()}-{date2.date()}.txt","w")
        print_file.write(f"""RMPL
        Attendance, Leave and Salary for the month ending {date2.day}.{date2.month}.{date2.year}""")
        print_file.close()