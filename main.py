from openpyxl import load_workbook
import os.path
import sys
import re
import textract
import subprocess
import fitz
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from io import StringIO
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from googleapiclient.discovery import build
import pprint
import io



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
            print(proc_asp.full_name)
            proc_asp.process_name()
            proc_asp.process_contacts()
            proc_asp.get_image()


class Processing(object):
    def __init__(self, object):
        self.full_name = object.name
        self.last_name = None
        self.first_name = None
        self.middle_name = None
        self.money = object.money
        self.text_resume = ''
        self.number = None
        self.email = None
        if object.resume != None:
            self.path = os.path.normpath(object.resume)
        # print(self.money)
        else:
            self.path = None


    def process_name(self):
        full_name = self.full_name.split()
        self.last_name = full_name[0]
        self.first_name = full_name[1]
        if len(full_name) > 2:
            self.middle_name = full_name[2]

    def process_contacts(self):
        if self.path != None:
            if self.path.endswith('.doc'):
                pp = pprint.PrettyPrinter(indent=4)
                SCOPES = ['https://www.googleapis.com/auth/drive']
                SERVICE_ACCOUNT_FILE = 'F:\hantflow_test\my-python-api-278708-2604d6b4e17e.json'

                credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
                service = build('drive', 'v3', credentials=credentials)

                results = service.files().list(
                    pageSize=100,
                    fields="nextPageToken, files(id, name, mimeType, parents, createdTime)",
                    q="name contains 'data'").execute()
                pp.pprint(results['files'])
                # path = os.path.normpath(r'F:\hantflow_test\task\Тестовое задание\Менеджер по продажам\Шорин Андрей.pdf')
                file_metadata = {'name': 'test',
                                 'mimeType': 'application/vnd.google-apps.document'
                                 }
                media = MediaFileUpload(self.path, mimetype='application/msword', resumable=True)
                file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                # File ID: 1eDjlhB-679FYneLYhECdWWTP0EdIidwm
                # file_mime = file.get('mimeType')
                file_id = file.get('id')
                # file_name = file.get('name')
                # print(file_id, file_mime, file_name)
                request = service.files().export_media(fileId=file_id,
                                                       mimeType='text/html'
                                                       )
                filename = f'{self.path[:-3]}html'
                fh = io.FileIO(filename, 'wb')
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
                    print("Download %d%%." % int(status.progress() * 100))
                # docx_path = self.path + 'x'
                # if not os.path.exists(docx_path):
                #     comand = f'antiword "{self.path}" > "{docx_path}"'
                #     os.system(comand)
                #     with open(docx_path) as f:
                #         text = f.read()
                #         # print(text)
                # else:
                #     # already a file with same name as doc exists having docx extension,
                #     # which means it is a different file, so we cant read it
                #     print('Info : file with same name of doc exists having docx extension, so we cant read it')
                #     text = ''
                # return text
            elif self.path.endswith('.pdf'):
                output_string = StringIO()
                with open(self.path, 'rb') as in_file:
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

    def get_image(self):
        if self.path != None:
            path = self.path
            if path.endswith('.pdf'):
                pdf = fitz.open(self.path)
                for i in range(len(pdf)):
                    for img in pdf.getPageImageList(i):
                        xref = img[0]
                        pix = fitz.Pixmap(pdf, xref)
                        if pix.n < 5:   # this is GRAY or RGB
                            pix.writePNG(f"{self.path[:-4]}_{i}{xref}.png")
                        else:  # CMYK: convert to RGB first
                            pix1 = fitz.Pixmap(fitz.csRGB, pix)
                            pix1.writePNG(f"{self.path[:-4]}_{i}{xref}.png")
                            pix1 = None
                        pix = None

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

