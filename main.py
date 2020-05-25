from openpyxl import load_workbook
import os.path
import sys

# test_base = openpyxl.load_workbook('./task/Тестовое задание/Тестовая база.xlsx')
# sheet = test_base.get_sheet_by_name('Лист1')
# print(sheet.cell(row=2, column=2).value)

class JS():
    def __init__(self, sheet, row):
        self.sheet = sheet
        self.row = row

    def search_resume(self):
        resume_path = os.path.realpath(f'{os.getcwd()}/{self.position}/{self.name}')
        print(resume_path)
        if os.path.exists(f'{resume_path}.doc') or os.path.exists(f'{resume_path}.pdf'):
            print('Да сучки')
        else:
            print('Nope')

    def job_seeker(self):
        self.position = self.sheet.cell(row=self.row, column=1).value
        if self.position != None:
            self.name = self.sheet.cell(row=self.row, column=2).value
            self.money = self.sheet.cell(row=self.row, column=3).value
            self.comment = self.sheet.cell(row=self.row, column=4).value
            self.status = self.sheet.cell(row=self.row, column=5).value
            print(self.position, self.name)
            self.search_resume()
        else:
            return sys.exit()




def open_base(path):
    test_base = load_workbook(path)
    return test_base['Лист1']

# def take_info(sheet):
#     while True:
#         row = +1
#         job_seeker(sheet, row)


if __name__ == '__main__':
    path = './task/Тестовое задание/Тестовая база.xlsx'
    new_path = os.path.basename(path)
    os.chdir(os.path.dirname(path))
    base = open_base(new_path)
    i = 2
    while True:
        js = JS(base, i)
        js.job_seeker()
        i += 1
