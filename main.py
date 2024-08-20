import os
import xlwings as xw
from datetime import datetime

# Indicate start of process
print("Start")

# Get Current directory
current_directory = os.getcwd()


##########################################################################################
# Opening of workbooks
##########################################################################################
first_book = xw.Book('reports.csv')
second_book = xw.Book('Disney Creative Scheduling.xlsx')
third_book = xw.Book('Disney_CreativeQA_Macro[1].xlsm')

# Access the sheets from the workbooks
scheduling_doc = second_book.sheets['FY24_Disney_Creative']
reports_sheet = first_book.sheets['reports']
macro_scheduling_doc = third_book.sheets['GOOGLE DOCS HERE']
macro_reports = third_book.sheets['CREATIVE CHECK DAILY RPT HERE']
macro_page = third_book.sheets['GENERATE REPORTS']

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

##########################################################################################
# Values loop
##########################################################################################

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


##########################################################################################
# Paste values from different sheets
##########################################################################################

macro_scheduling_doc.range('A1').value = campaign_values
macro_scheduling_doc.range('B1').value = adconcept_values
macro_scheduling_doc.range('E1').value = creative_doc_values
macro_scheduling_doc.range('F1').value = start_values
macro_scheduling_doc.range('G1').value = end_values

macro_reports.range('A1').value = date_values
macro_reports.range('B1').value = creative_values
macro_reports.range('C1').value = creative_id_values
macro_reports.range('D1').value = total_ads_values

# Change Directory to current directory
macro_page.range('B10').value = current_directory
macro_page.range('I10').value = "Pass 1"

##########################################################################################
# First macro
##########################################################################################

print("Start Macro")

try:

    # Macro Run
    macro = third_book.macro('Module1.GenerateDisneyReports')
    macro()
    print("Running first macro")

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    try:
        print("Finished Running first macro")
    except:
        pass  
    
    
##########################################################################################
# Replacement of MAUI Codes Start
##########################################################################################


search_items = ['WDWDOM', 'WDWEPCOT', 'WDWFLRES', 'WDWRSTS', 'WDWDHS', 'DSPNGS', 'WDWRES', 'WDWEEC', 'WDWLATAM', 'CONSUMER', '320x50', '0', 'DIQF']
print("Replacing MAUI CODES")
search_col = macro_scheduling_doc.range('C:C')

# Iterate through the column values
for i, cell_value in enumerate(search_col.value):
    if cell_value in search_items:
        # Copy command
        col_e_value = macro_scheduling_doc.range(f'E{i+1}').value
        
        # Split Command
        split_values = col_e_value.split('_')
        if len(split_values) >= 4: 
            new_value = split_values[3] #Number 3 kasi pang apat yung maui code sa array
        else:
            new_value = ''  # Or handle cases where there are fewer than 4 items

        # Update column C with the 4th element
        macro_scheduling_doc.range(f'C{i+1}').value = new_value
        
# File path and File name
now = datetime.now()
formatted_date = now.strftime('%m.%d.%Y')

##########################################################################################
# Second macro
##########################################################################################


print(f'Current Date: {formatted_date}')

macro_page.range('I10').value = formatted_date + "_Creative_QA_Report"

try:

    # Macro Run
    macro = third_book.macro('Module1.GenerateDisneyReports')
    macro()
    print("Running Second macro")

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    try:
        print("Finished Running Second macro")
    except:
        pass  
    
    
##########################################################################################
# Cleaning the report
##########################################################################################

print("Cleaning the report")
def rgb_to_excel_color(r, g, b):
    return (r << 16) + (g << 8) + b

r, g, b = 0, 255, 255 
color_rgb = rgb_to_excel_color(r, g, b)

output_report = xw.Book(formatted_date + "_Creative_QA_Report.xlsx")

# Define the range to search within (e.g., the used range of the sheet)

RemoveRotation = output_report.sheets['Remove From Rotation']
ManualChecking = output_report.sheets['Manual Checking Needed']


RemoveRotation.range('A1').value = 'There are no creatives (having a minimum of 150 impressions) in rotation past their end dates.'
ManualChecking.range('A1').value = 'The below creatives have multiple end dates associated with the given MAUI code and job number.'

used_range = RemoveRotation.range('G3').current_region 
# Iterate over each cell in the range
for cell in used_range:
    if isinstance(cell.value, (int, float)) and cell.value >= 150:
        cell.api.Interior.Color = color_rgb 


##########################################################################################
# Finishing up
##########################################################################################
first_book.close()
second_book.close()

# Save and close the target workbook
output_report.save()

#close all workbook
output_report.close()

print("End")