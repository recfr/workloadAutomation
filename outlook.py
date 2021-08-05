import os
import win32com.client as client
from PIL import ImageGrab
import datetime
from pathlib import Path


class EmailSender:
    def __init__(self):
        self.desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        self.employeeList = Path('employees.txt').read_text().replace('\n', '; ')
        self.carbonCopyList = Path('carboncopy.txt').read_text()

    def createNewMail(self, workBook_path):
        calendar_week = datetime.date.today().isocalendar()[1]
        outlook = client.Dispatch('outlook.application')

        html_body = "<html>" \
               "<head></head>" \
               "<body>" \
               "<p>Merhaba Arkadaşlar,<br></p>" \
               "<p>RMLS tarihi gelen iş yüklerini bugün tamamlamanız ricasıyla, iyi çalışmalar dilerim.</p>" \
               "<div>image_source(graphic)" \
               "<img src{}></img>" \
               "</div>" \
               "<p><u><b>ANALİZ:</b></u></p>" \
               "<p><ol></p>" \
               "<li> KW%(cw_early)i ve KW%(cw_today)i genel olarak çalışma kapasitemize uygun çalışılmıştır.</li>" \
               "<li> Şu an iş yükünde “ ” Fehler, “ “ Açık Kapsam,  “ “ RMLS,  “ “ AKT ve  “ “ Bildiri bulunmaktadır.</li>" \
               "</ol>" \
               "<div>image_source(pivot_table)" \
               "<img src{}></img>" \
               "</div>" \
               "<b><p>ÖNEMLİ : </p></b> " \
               "<b>“Açık Kapsam: Kontrol Edilecek”</b> sütununda 47’den dönen, geçmiş tarihli, satış nachtragı ya da başka herhangi" \
               "sebepten açılan kapsamlar olabilir. Ekteki excel dokümanında “Sheet1” isimli sayfada, “Çalışma Tipi” sütununda " \
               "“Kontrol Edilecek” olarak belirtilen çalışmaları kontrol etmenizi rica ederim. " \
               "<b><p>NOT: Lütfen yetişemeyecek kapsamlar hakkında koordinatörümüzü bilgilendirin. </p></b>" \
               "</body>" \
               "</html> " % {"cw_early": int(calendar_week) - 1,
                             "cw_today": int(calendar_week)
                             }

        # excel = client.Dispatch('Excel.application')
        # wb = excel.Workbooks.Open(workBook_path)
        # sheet = wb.Sheets['Pivot Table']
        # excel.visible = 1
        # copyrange = sheet.Range('A1:J13')
        # copyrange.CopyPicture(Appearance=1, Format=2)
        # ImageGrab.grabclipboard().save('paste.png')

        mail = outlook.CreateItem(0)
        mail.To = self.employeeList
        mail.CC = self.carbonCopyList

        mail.Subject = 'İş Yükü'
        mail.HTMLBody = html_body

        # attachment = self.desktop + "\\" + workBook_path
        # mail.Attachments.Add(attachment)

        # mail.Send()
        mail.Display()