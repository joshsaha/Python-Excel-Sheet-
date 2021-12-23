import xlsxwriter
import xlrd

filename = 'DBS_GPI_Music-OCT23-2018-2HzTapSelfMarked-Event.xlsx'
workbook = xlsxwriter.Workbook(filename)
workbook2 = xlrd.open_workbook(filename)

worksheet = workbook2.sheet_by_name('Sheet1')
worksheet2 = workbook.add_worksheet()

for i in range(0, 101):
    for j in range(0, 6):
        worksheet2.write(i, j, (float) (format(worksheet.cell(i, j).value)))

for i in range(0, 101):
    a = (float) (format(worksheet.cell(i+1, 5).value))
    b = (float) (format(worksheet.cell(i, 5).value))
    SEsample = a - b
    SEsampleminus = b - a
    SEsamples = SEsample/5000
    SEerror = 1 - SEsamples
    # starts from 7 ends at 10
    worksheet2.write(i, 7, SEsample)
    worksheet2.write(i, 8, SEsamples)
##  worksheet2.write(i, 9, SEsampleminus)
##  worksheet2.write(i, 10, SEerror)
    
workbook.close()
