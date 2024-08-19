import os
import xlwings as xw
from datetime import datetime

first_book = xw.Book('reports.csv')
second_book = xw.Book('Disney Creative Scheduling.xlsx')
third_book = xw.Book('macro.xlsm')

# Access the sheets from the workbooks
scheduling_doc = second_book.sheets['FY24_Disney_Creative']
reports_sheet = first_book.sheets['reports']
macro_scheduling_doc = third_book.sheets['GOOGLE DOCS HERE']
macro_reports = third_book.sheets['CREATIVE CHECK DAILY RPT HERE']

search_items = ['WDWRES', 'ITEM2', 'ITEM3']  # Replace with your items

search_col = macro_scheduling_doc.range('C:C')

# Get the current date
now = datetime.now()
formatted_date = now.strftime('%m.%d.%Y')


print(f'Current Date: {formatted_date}')

current_directory = os.getcwd()
macro_page = third_book.sheets['GENERATE REPORTS']
macro_page.range('B10').value = current_directory
macro_page.range('I10').value = formatted_date + "_Creative_QA_Report"


