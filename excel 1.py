# import module 
import openpyxl 
  
# load excel with its path 
wrkbk = openpyxl.load_workbook("project edited py1.xlsx") 
  
sh = wrkbk.active
one=["Male", "Under 20", "Islam", "Hausa", "Never"", Not at all", "Not satisfied at all", "Yes"]
  
# iterate through excel and display data 
for row in sh.iter_rows(min_row=1, min_col=1, max_row=12, max_col=3): 
    for cell in row:
        if cell in one:
            cell.value = "1"
        else:
            pass
        #print(cell.value, end=" ") 
    #print() 
wrkbk.save("project edited py1.xlsx")
