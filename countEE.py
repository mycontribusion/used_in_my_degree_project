import openpyxl
import random

# Specify the Excel file path to be edited
excel_file_path = 'projectwith2age111.xlsx'

# Load the existing Excel workbook
workbook = openpyxl.load_workbook(excel_file_path)

# Select the sheet you want to edit (assuming the first sheet for simplicity)
sheet = workbook.worksheets[0]

low=[1,"Low"]
moderate=[2,"Moderate"]
high=[3,"High"]

low_count=[]
moderate_count=[]
high_count=[]

for column in sheet.iter_cols():
    column_name = column[0].value
    #print(column_name)
    if column_name == "Burnout Level":
        print(column_name)
        for cell in column:
            if cell.value in low:
                low_count.append(cell.value)
                #print(low_count)
            elif cell.value in moderate:
                moderate_count.append(cell.value)
                #print(moderate_count)
            elif cell.value in high:
                high_count.append(cell.value)

low_len=len(low_count)/192
low_perc=low_len*100
    
print(len(low_count))
#print(low_len)
print(low_perc)

moderate_len=len(moderate_count)/192
moderate_perc=moderate_len*100
    
#print(moderate_count)
print(len(moderate_count))
print(moderate_perc)

high_len=len(high_count)/192
high_perc=high_len*100
    
#print(high_count)
print(len(high_count))
print(high_perc)






#print(moderate_count)
#print(len(moderate_count))
#print(high_count)
#print(len(high_count))
