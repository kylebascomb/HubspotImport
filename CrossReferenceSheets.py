import pandas as pd
import xlrd as xlrd
from Company import Company
from tempfile import TemporaryFile
from xlwt import Workbook


def format_domain(domain):
    length = len(domain)
    if length >= 12:
        if "http://" == domain[0:7]:
            domain = domain[7:]
            domain = domain.lower()
            return domain
        else:
            return "format_error"
    else:
        return "format_error"


def write_to_file(newNames, exportSheet):
    rowIndex = 1
    ORIGIN = "John Hiller List"
    STAGE = "Lead"
    OWNER = "Pia Corpuz"
    #Write headers
    exportSheet.write(0, 0, "Name")
    exportSheet.write(0, 1, "Company Domain")
    exportSheet.write(0, 2, "Origin")
    exportSheet.write(0, 3, "Lifecycle Stage")
    exportSheet.write(0, 4, "Owner")
    #Write All Data
    for company in newNames:
        exportSheet.write(rowIndex, 0, company.name)
        exportSheet.write(rowIndex, 1, company.domain)
        exportSheet.write(rowIndex, 2, ORIGIN)
        exportSheet.write(rowIndex, 3, STAGE)
        exportSheet.write(rowIndex, 4, OWNER)
        rowIndex += 1

def find_col(keyword, sheet):
    HEADER_ROW = 0
    for col in range(sheet.ncols):
        if sheet.cell(HEADER_ROW, col).value == keyword:
            return col
    return -1


#sets
dbNames = set()
refNames = set()
#import books
DB_NAME = "database.xlsx"
REF_NAME = "reference.xlsx"
book1 = xlrd.open_workbook(DB_NAME)
book2 = xlrd.open_workbook(REF_NAME)
sheet1 = book1.sheet_by_index(0)
sheet2 = book2.sheet_by_index(0)
#export books
exportBook = Workbook()
exportSheet = exportBook.add_sheet('Sheet 1')
#find the index of each column of interest
colNameDB = find_col("Name", sheet1)
colDomainDB = find_col("Company Domain Name", sheet1)
colNameRF = find_col("Company Name", sheet2)
colDomainRF = find_col("Company Website", sheet2)
#fill each set with the correct information
for row in range(sheet1.nrows):
    dbNames.add(Company(sheet1.cell(row, colNameDB).value, sheet1.cell(row, colDomainDB).value))

for row in range(sheet2.nrows):
    domain = sheet2.cell(row, colDomainRF).value
    if format_domain(domain) != "format_error":
        domain = format_domain(domain)
        refNames.add(Company(sheet2.cell(row, colNameRF).value, domain))

#create variables to store the new data
newNames = refNames.difference(dbNames)

if len(newNames) == 0:
    print("Nothing to update")
else:
    #write info to file
    write_to_file(newNames, exportSheet)
    exportBook.save('export.xls')
    print("Export successful")




