# Description
 This is an excel automation program developed for private use. 
###### How to use ?

- Name of input excel file is not important you just need to set the name of worksheet **Sheet1**

- Columns' order must be as follows.
1. SollRückmeldetermin Leitstand
2. Konstruktionstermin Soll
3. Dokument
4. Beschreibung
5. BB-Nummer
6. Dokumentstatus
7. Auftragsphase
8. BB Beschreibung
9. KSW-Status
10. Dokumentennummer Maßnahme
11. Sachbearbeiter

- Click to "**Execute**" button to create an edited excel workbook.
- Then "**Create an E-Mail**" button will be clickable to create a new E-mail.

###### Information
employees.txt   : This file should contain your reciever list.
carboncopy.txt  : This file should contain your CC (CarbonCopy) list.

## Libs 
- [Pandas](https://pandas.pydata.org/)
- [PyQt5](https://pypi.org/project/PyQt5/)
- [Datetime](https://docs.python.org/3/library/datetime.html)
- [xlsxwriter](https://xlsxwriter.readthedocs.io/)
- [pyOutlook](https://pypi.org/project/pyOutlook/)


