from openpyxl import load_workbook
import os.path
import sys
import re
import textract
import subprocess
import PyPDF4
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from io import StringIO
from datetime import datetime
import calendar


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
            # print(type(self.resume))
        elif os.path.exists(f'{resume_path}.pdf'):
            self.resume = os.path.abspath(f'{resume_path}.pdf')
            # print(type(self.resume))
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
            proc_asp.process_contacts(aspirant)


class Processing(object):
    def __init__(self, object):
        self.last_name = None
        self.first_name = None
        self.middle_name = None
        self.money = object.money
        self.text_resume = ''
        self.number = None
        self.email = None
        # print(self.money)


    def process_name(self,object):
        full_name = object.name.split()
        self.last_name = full_name[0]
        self.first_name = full_name[1]
        if len(full_name) > 2:
            middle_name = full_name[2]

    def process_contacts(self, object):
        path = os.path.normpath(object.resume)
        soffice = os.path.normpath('"C:\Program Files (x86)\OpenOffice 4\program\soffice.exe"')
        if path.endswith('.doc'):
            prog = subprocess.Popen(['runas', '/user:Admin', soffice, '--headless', '--convert-to', 'docx', path])
            prog.communicate()
            print(1)
            # path_split = os.path.split(path)
            # path_head = path_split[0]
            # path_tail = path_split[1].split()
            # new_name = ''
            # for i in path_tail:
            #     new_name = new_name + i
            # path_new = os.path.join(path_head, new_name)
            #
            # copy_str = f'copy "{path}" "{path_new}"'
            # print(copy_str)
            # os.system(copy_str)
            # os.system('copy "' + path + '" ' + path_new)
            # str = textract.process(path_new)
            # email_regezp = r"^[A-Za-z0-9\.\+_-]+@[A-Za-z0-9\._-]+\.[a-zA-Z]*$"
            # email_search = re.findall(email_regezp, str)
            # print(email_search)
            # antiword_str = f'antiword "{path}" > "{path}x"'
            # print(antiword_str)
            # os.system(antiword_str)
            # with open(path_x) as f:
            #     text = f.read()
            #     print(text)
        elif path.endswith('.pdf'):
            output_string = StringIO()
            with open(path, 'rb') as in_file:
                parser = PDFParser(in_file)
                doc = PDFDocument(parser)
                rsrcmgr = PDFResourceManager()
                device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
                interpreter = PDFPageInterpreter(rsrcmgr, device)
                for page in PDFPage.create_pages(doc):
                    interpreter.process_page(page)
            self.text_resume = output_string.getvalue()
            self.number = re.search(r'(\+7|8).*?(\d{3}).*?(\d{3}).*?(\d{2}).*?(\d{2})', self.text_resume)
            regex = re.compile(("([a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`"
                                "{|}~-]+)*(@|\sat\s)(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(\.|"
                                "\sdot\s))+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)"))
            self.email = re.search(regex, self.text_resume)
            full_date = re.search(r'(\d{2}(\-| |\/)(\d{2}|января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)(\-| |\/)\d{4})', self.text_resume)
            if full_date != None:
                date = full_date.group().split()
                self.birthday_day = date[0]
                RU_MONTH_VALUES = {
                    'января': 1,
                    'февраля': 2,
                    'марта': 3,
                    'апреля': 4,
                    'мая': 5,
                    'июня': 6,
                    'июля': 7,
                    'августа': 8,
                    'сентября': 9,
                    'октября': 10,
                    'ноября': 11,
                    'декабря': 12,}
                month = date[1]
                self.birthday_month = RU_MONTH_VALUES.get(month)
                self.birthday_year = date[2]

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

