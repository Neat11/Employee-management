import datetime
from openpyxl import load_workbook
from destination import path
class EmployeeAdd():

    def get_maximum_rows(sheet_object):
        rows = 0
        for max_row, row in enumerate(sheet_object, 1):
            if not all(col.value is None for col in row):
                rows += 1
        return rows


    def add_emp(self, values):
        wb = load_workbook(path)
        sheet = wb["MAST1"]
        rows = self.get_maximum_rows(sheet)
        char = 'A'
        try:
            date_joined = values["DTJOIN"].split("-")
            values["DTJOIN"] = datetime.datetime(int(date_joined[2]),int(date_joined[1]),int(date_joined[0]))
            leave_init = values["LEAVE_INIT"].split("-")
            values["LEAVE_INIT"] = datetime.datetime(int(leave_init[2]),int(leave_init[1]),int(leave_init[0]))
            for i in range(3,13):
                print(values[list(values.keys())[i]])
                if (values[list(values.keys())[i]].isnumeric()) or values[list(values.keys())[i]]=="":
                    continue
                else:
                    return "There are One or more Invalid inputs"
            print([values[i] for i in values])
            for i in values:
                if (values[i] == values["DTJOIN"]) or (values[i] != values["LEAVE_INIT"]):
                    sheet[char+str(rows+1)] = values[i]
                elif (values[i].isnumeric()):
                    sheet[char+str(rows+1)] = int(values[i])
                elif values[i] == "TRUE":
                    sheet[char+str(rows+1)] = True
                elif values[i] == "FALSE":
                    sheet[char+str(rows+1)] = False
                elif values[i] == "":
                    sheet[char+str(rows+1)] = 0
                else:
                    sheet[char+str(rows+1)] = values[i]
                char = chr(ord(char)+1)
            wb.save(path)
            return "Employee added successfully"
        except Exception as e:
            return str(e)
