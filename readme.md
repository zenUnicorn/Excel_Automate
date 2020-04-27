#Automate your excel SpreadSheet using python and perform repetitive borinng jobs in seconds
import openpyxl as xl

from openpyxl.chart import BarChart, Reference

#processing thousands of spreadsheet
#automate repetitive boring stuff that waste your time
def process_excel(filename):
    wb = xl.load_workbook(filename)
    #getting the sheet in the excel spreadsheet
    sheet = wb['Sheet1']
    #to access a cell use
    # cell = sheet['a1'] or 
    # cell  = sheet.cell(1,1) #row , col
    # print(cell.value)

    #getting the numbers of rows in th sheet
    # print(sheet.max_row)

    #looping thru the excel file
    # for row in range(1, sheet.max_row+1):
    #     print(row)
    #getting the values in the third column
    for row in range(2, sheet.max_row+1):
        cell3 = sheet.cell(row, 3) #row 3
        corrected_price = cell3.value * 0.9
        corrected_price_cell = sheet.cell(row, 4)
        corrected_price_cell.value = corrected_price

    #referencing the spreadsheet to create a chart, first we get the lowset and hughest row
    values = Reference(sheet, 
                min_row=2, 
                max_row=sheet.max_row,
                min_col=4,
                max_col=4)
    chart = BarChart()
    chart.add_data(values)
    #adding the chart to the spreadsheet,to the coordinate e2
    sheet.add_chart(chart, 'e2')

    #then save when youre done
    wb.save(filename)

