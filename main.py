import xlwings as xw
import csv

# Indicate start of process
print("Start")

# Open the source and target workbooks
first_book = xw.Book('reports.csv')
second_book = xw.Book('Disney Creative Scheduling.xlsx')
third_book = xw.Book('macro.xlsm')

# Access the sheets from the workbooks
scheduling_doc = second_book.sheets['FY24_Disney_Creative']
reports_sheet = first_book.sheets['reports']
macro_scheduling_doc = third_book.sheets['GOOGLE DOCS HERE']
macro_reports = third_book.sheets['CREATIVE CHECK DAILY RPT HERE']


# Scheduling Docs Data
campaign_range = scheduling_doc.range('A:A').value
adConcept_range = scheduling_doc.range('B:B').value
creative_doc_range = scheduling_doc.range('E:E').value
start_range = scheduling_doc.range('G:G').value
end_range = scheduling_doc.range('H:H').value

# Reports Data
date_range = reports_sheet.range('A:A').value
creative_range = reports_sheet.range('B:B').value
creative_id_range = reports_sheet.range('C:C').value
totalads_range = reports_sheet.range('D:D').value

# First sheet paste
campaign_values = [[item] for item in campaign_range if item is not None]
adconcept_values = [[item] for item in adConcept_range if item is not None]
creative_doc_values = [[item] for item in creative_doc_range if item is not None]
start_values = [[item] for item in start_range if item is not None]
end_values = [[item] for item in end_range if item is not None]

# Second sheet paste
date_values = [[item] for item in date_range if item is not None]
creative_values = [[item] for item in creative_range if item is not None]
creative_id_values = [[item] for item in creative_id_range if item is not None]
total_ads_values = [[item] for item in totalads_range if item is not None]


# eto yung nag papaste
macro_scheduling_doc.range('A1').value = campaign_values
macro_scheduling_doc.range('B1').value = adconcept_values
macro_scheduling_doc.range('E1').value = creative_doc_values
macro_scheduling_doc.range('F1').value = start_values
macro_scheduling_doc.range('G1').value = end_values

macro_reports.range('A1').value = date_values
macro_reports.range('B1').value = creative_values
macro_reports.range('C1').value = creative_id_values
macro_reports.range('D1').value = total_ads_values


print("Start Macro")

try:

    # Macro Run
    macro = third_book.macro('Module1.GenerateDisneyReports')
    macro()
    print("Running first macro")

    print("Finished first macro")

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    try:
        print("Finished Running first macro")
    except:
        pass  

find_maui = 'WDWRES'

search_col = macro_scheduling_doc.range('C:C')

found_rows = []

# Iterate through the column values
for i, cell_value in enumerate(search_col.value):
    if cell_value == find_maui:
        row_range = macro_scheduling_doc.range(f'C{i+1}:E{i+1}')  
        row_values = row_range.value
        found_rows.append(row_values)

# Output the found rows
if found_rows:
    print(f'Found "{find_maui}" in the following rows:')
    for row in found_rows:
        print(row)
else:
    print(f'No "{find_maui}" found')
    
# print(found_rows)

print("End")