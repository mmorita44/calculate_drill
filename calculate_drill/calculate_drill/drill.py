import random
import openpyxl
import traceback

class Drill:

    def __init__(self, file):
        self.row_max = 27
        self.cloumn_max = 20
        self.file = file

    def calculate(self):
        try:
            book = openpyxl.load_workbook(self.file)
            sheet = book['Sheet1']
            records = []
            for first in range(100, 999):
                for second in range(10, 99):
                    records.append([first, second])
            random.shuffle(records)

            row_pos = 1
            column_pos = 1
            for record in records:
                if row_pos >= self.row_max:
                    row_pos = 1
                    column_pos += 5

                if column_pos >= self.cloumn_max:
                    break

                sheet.cell(row=row_pos, column=column_pos, value=record[0])
                sheet.cell(row=row_pos, column=column_pos + 2, value=record[1])
                row_pos += 2

            book.save(self.file)
            return 0
        except Exception as ignored:
            print(traceback.format_exc())
            return 1
