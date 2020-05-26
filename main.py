from openpyxl import load_workbook
import os.path
import sys


class JS:

    def __init__(self, position, name, money, comment, status):
        self.position = position
        self.name = name
        self.money = money
        self.comment = comment
        self.status = status
        self.resume = None

    def search_resume(self):
        resume_path = os.path.realpath(f'{os.getcwd()}/{self.position}/{self.name}')
        # print(resume_path)
        if os.path.exists(f'{resume_path}.doc'):
            self.resume = os.path.abspath(f'{resume_path}.doc')
            print(self.resume)
        elif os.path.exists(f'{resume_path}.pdf'):
            self.resume = os.path.abspath(f'{resume_path}.pdf')
            print(self.resume)
        else:
            print('Nope')


class Base:

    def __init__(self, path):
        self.path = path
        self.base = load_workbook(self.path)
        self.all_job_seekers = []

    def all_job_seeker(self):
        all_sheet = self.base.sheetnames
        for i in all_sheet:
            current_sheet = self.base[i]
            row = 2
            # all_job_seeker = []
            while True:
                position = current_sheet.cell(row=row, column=1).value
                if position != None:

                    new_job_seeker = [position,
                                      current_sheet.cell(row=row, column=2).value,
                                      current_sheet.cell(row=row, column=3).value,
                                      current_sheet.cell(row=row, column=4).value,
                                      current_sheet.cell(row=row, column=5).value]
                    self.all_job_seekers.append(new_job_seeker)
                    row += 1
                else:
                    # print(self.all_job_seekers)
                    return False

    def new_job_seeker(self):
        for i in self.all_job_seekers:
            aspirant = JS(*i)
            aspirant.search_resume()
            proc_asp = Processing(aspirant)
            # print(proc_asp.last_name)


class Processing(object):
    def __init__(self, object):
        self.last_name = None
        self.first_name = None
        self.middle_name = None
        self.money = object.money

        print(self.money)


    def process_name(self,object):
        full_name = object.name.split()
        self.last_name = full_name[0]
        self.first_name = full_name[1]
        if len(full_name) > 2:
            middle_name = full_name[2]

    def process_contacts(self, object):
        path = object.resume




class New_JS(object):

    def create_json(self):
        pass




if __name__ == '__main__':
    path = './task/Тестовое задание/Тестовая база.xlsx'
    new_path = os.path.basename(path)
    os.chdir(os.path.dirname(path))
    base = Base(new_path)
    base.all_job_seeker()
    base.new_job_seeker()

