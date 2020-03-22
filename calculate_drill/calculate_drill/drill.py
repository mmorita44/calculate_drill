import random
import openpyxl
import traceback

class Drill:

    def __init__(self):
        self.row_max = 27
        self.cloumn_max = 36
        self.file = 'document/足し算.xlsx'

    def calculate(self):
        try:
            book = openpyxl.load_workbook(self.file)
            sheet = book['Sheet1']
            records = []
            for first in range(1, 10):
                for second in range(1, 10):
                    for third in range(1, 10):
                        records.append([first, second, third])
            random.shuffle(records)

            row_pos = 1
            column_pos = 1
            for record in records:
                if row_pos >= self.row_max:
                    row_pos = 1
                    column_pos += 7

                if column_pos >= self.cloumn_max:
                    break

                sheet.cell(row=row_pos, column=column_pos, value=record[0])
                sheet.cell(row=row_pos, column=column_pos + 2, value=record[1])
                sheet.cell(row=row_pos, column=column_pos + 4, value=record[2])
                row_pos += 2

            book.save(self.file)
            return 0
        except Exception as ignored:
            print(traceback.format_exc())
            return 1
