import openpyxl
import os

#load the PDM and Arena workbooks
#make sure the workbook is named 'PDM.xlsx', and 'Arena.xlsx'

wb_pdm = openpyxl.load_workbook('PDM.xlsx')
wb_arena = openpyxl.load_workbook('Arena.xlsx')

#identify first sheets to be used
sheet_pdm = wb_pdm['Sheet1']
sheet_arena = wb_arena['Sheet1']

#number of rows and last row in PDM and Arena
num_rows_pdm = sheet_pdm.max_row
num_rows_arena = sheet_arena.max_row

#initialize table for matching number build out
matching_table = []
table_index = 0
i= 0

#this creates a table of all components that are in PDM and Arena, with various name parsing
for i in range(1,num_rows_pdm):
    current_pdm = 'A' + str(i)
    current_value_ext = sheet_pdm['%s' %current_pdm].value
    current_value = os.path.splitext(current_value_ext)[0] #filename minus extension
    current_value_last = current_value[-11:] #last 11 digits of filename for incorrectly named files

    for x in range(1,num_rows_arena):
        current_arena = 'A' + str(x)
        current_arena2 = sheet_arena['%s' %current_arena].value
        if current_value == current_arena2:
            matching_table.append(current_value)
            table_index = table_index + 1
        elif current_value_last == current_arena2:
            matching_table.append(current_value_last)
            table_index+=1

print(matching_table)

# for item in matching_table:
#
#
#
