import os
import xlwings as xw
from datetime import datetime

# Access the sheets from the workbooks

now = datetime.now()
formatted_date = now.strftime('%m.%d.%Y')

def rgb_to_excel_color(r, g, b):
    return (r << 16) + (g << 8) + b

r, g, b = 0, 255, 255 
color_rgb = rgb_to_excel_color(r, g, b)

output_report = xw.Book(formatted_date + "_Creative_QA_Report.xlsx")

# Define the range to search within (e.g., the used range of the sheet)

reports_sheet = output_report.sheets['Remove From Rotation']
used_range = reports_sheet.range('G3').current_region 

# Iterate over each cell in the range
for cell in used_range:
    if isinstance(cell.value, (int, float)) and cell.value >= 150:
        cell.api.Interior.Color = color_rgb 
