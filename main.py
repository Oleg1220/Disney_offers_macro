import xlwings as xw

print("Start")

try:
    wb = xw.Book('macro.xlsm') 
    wb.Open()
    
    # Macro Run
    print("Running macro")
    macro = wb.macro('Module1.GenerateDisneyReports')
    macro()  # Call the macro

    print("Finished")

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    try:
        wb.close()
    except:
        pass  

    try:
        xw.apps.active.quit()  # Quit the Excel application
    except:
        pass  
