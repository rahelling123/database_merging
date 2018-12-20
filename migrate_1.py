import openpyxl
import os
import numpy

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
master_data = ['original file name', 'name', 'pdm_revision', 'pdm_state', 'revised by', 'arena_revision', 'item_phase']
arena_data = ['number', 'name', 'revision', 'phase', 'owner']
arena_data_add = []
table_index = 0

#create array of all Arena files
for i in range(2,num_rows_arena):
    arena_data_add = []
    arena_data_add.append(sheet_arena[('A'+str(i))].value)
    arena_data_add.append(sheet_arena[('B' + str(i))].value)
    arena_data_add.append(sheet_arena[('C' + str(i))].value)
    arena_data_add.append(sheet_arena[('D' + str(i))].value)
    arena_data_add.append(sheet_arena[('E' + str(i))].value)
    arena_data=numpy.vstack((arena_data,arena_data_add))


#this creates a table of all components that match in PDM and Arena, with various name parsing
for i in range(1,num_rows_pdm):
    current_pdm = 'A' + str(i)
    current_value_ext = sheet_pdm['%s' %current_pdm].value
    current_value = os.path.splitext(current_value_ext)[0] #filename minus extension
    current_value_last = current_value[-11:] #last 11 digits of filename for incorrectly named files

    for x in range(1,num_rows_arena):
        match_data=[]
        current_arena = 'A' + str(x)
        current_arena = sheet_arena['%s' %current_arena].value
        if current_value == current_arena:
            match_data.append(current_value_ext)#start building row of data for matches
            match_data.append(current_value)
            match_data.append(sheet_pdm[('C' + str(i))].value) # revision pdm
            match_data.append(sheet_pdm[('D' + str(i))].value) # state pdm
            match_data.append(sheet_pdm[('E' + str(i))].value) # revised by pdm
            match_data.append(sheet_arena[('C' + str(i))].value) #revision arena
            match_data.append(sheet_arena[('D' + str(i))].value) #phase state arena
            master_data = numpy.vstack((master_data,match_data))
            matching_table.append(current_value)
            table_index = table_index + 1
        elif current_value_last == current_arena:
            match_data.append(current_value_ext)#start building row of data for matches
            match_data.append(current_value_last)
            match_data.append(sheet_pdm[('C' + str(i))].value) # revision pdm
            match_data.append(sheet_pdm[('D' + str(i))].value) # state pdm
            match_data.append(sheet_pdm[('E' + str(i))].value) # revised by pdm
            match_data.append(sheet_arena[('C' + str(i))].value) #revision arena
            match_data.append(sheet_arena[('D' + str(i))].value) #phase state arena
            master_data = numpy.vstack((master_data,match_data))
            matching_table.append(current_value_last)
            table_index+=1

# print(matching_table)
print(master_data[1])

new_wb = openpyxl.Workbook()
sheet = new_wb.active
master_data[1]

for i in range(1,len(master_data)):
     sheet['A' + str(i)] = master_data[i,1]
     sheet['B' + str(i)] = master_data[i,2]
     sheet['C' + str(i)] = master_data[i,3]
     sheet['D' + str(i)] = master_data[i,4]
     sheet['E' + str(i)] = master_data[i,5]
     sheet['F' + str(i)] = master_data[i,6]

new_wb.save('first_excel_output.xlsx')



#initialize the vari