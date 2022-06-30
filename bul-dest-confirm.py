#take the vrc barcodes from the destruction report for Brig Eastman and 
#return any vrc barcodes that are not on the Total Inventory List - Main.xlsx file
#IMPORTANT - this script assumes that the data in question is located on the first sheet of each excel workbook accessed

#get pandas to read Sheet1 of all excel files stored in the file_names list:
import math
import pandas
file_names = [r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Brig Eastman.xlsx', 
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main.xlsm']
file_name_dict = {"file_name" + str(i+1): file_names[i] for i in list(range(0,len(file_names)))}
sheet_name_dict = {"file_name" + str(i+1): list(pandas.ExcelFile(file_name_dict["file_name" + str(i+1)]).sheet_names) for i in list(range(0,len(file_names)))}
worksheet_dict = {"file_name" + str(i+1): pandas.read_excel(file_name_dict[list(file_name_dict.keys())[i]], sheet_name = sheet_name_dict['file_name' + str(i+1)][0]) for i in list(range(0,len(file_names)))}
rows_columns_dict = {"file_name" + str(i+1): worksheet_dict[list(worksheet_dict.keys())[i]].shape for i in list(range(0,len(file_names)))}
headers_dict = {"file_name" + str(i+1): list(worksheet_dict['file_name' + str(i+1)].columns) for i in list(range(0,len(file_names)))}

# create a dictionary of a dictionary of lists representing each separate excel worksheet and its corresponding column data:
#add the column data (excluding headers bc worksheet.iloc(0) represents the second row of data)
#to each dictionary key representing each column of "Sheet1" of each excel file

letter_string = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
alpha_list = list(letter_string)
Alpha_list = []
factors = list(range(1,28))
alphanum = len(alpha_list)
while len(Alpha_list) < alphanum*factors[-1]:
	for letter in alpha_list:
		if len(Alpha_list) < alphanum:
			Alpha_list.append(letter)
		else:
			if ((factors[alpha_list.index(letter)])*(alphanum)) <= len(Alpha_list) < ((factors[alpha_list.index(letter)]+1)*(alphanum)):
				for second_letter in alpha_list:
					Alpha_list.append(letter + second_letter)

alphanumeric_dict = {letter: index for index, letter in list(enumerate(Alpha_list))}
column_dict = {file: {} for file in file_name_dict.keys()}
for file in file_name_dict.keys():
	alpha_labels = [Alpha_list[i] for i in list(range(0, len(worksheet_dict["file_name" + str(list(file_name_dict.keys()).index(file) + 1)].columns)))]
	headers = worksheet_dict["file_name" + str(list(file_name_dict.keys()).index(file) + 1)].columns
	for alpha_label in alpha_labels:
		column_dict[file][alpha_label] = []
		column_dict[file][alpha_label].append(headers[alpha_labels.index(alpha_label)])
	for col_data_list in list(column_dict[file].keys()):
		for cell in list(worksheet_dict[file].iloc[:,alphanumeric_dict[list(column_dict[file].keys())[list(column_dict[file].keys()).index(col_data_list)]]]):
			column_dict[file][col_data_list].append(cell)

# #define function that returns cell at intersection of row X and column Y example: cell(2,'A')

def cell(file_name,column,row):
	return print(column_dict[file_name][column][row-1])

#create query function for data on single sheet (Use this to query Total Inventory List - Main.xlsm)

def query(file_name, col_label, filter_file_name = file_name_dict, filter_col_label, filter_criterion):
	query_list = [cell for cell in column_dict[file_name][col_label] if column_dict[file_name][filter_col_label][list(column_dict[file_name][col_label]).index(cell)] == filter_criterion]
	if query_list == []:
		return "None"
	else:
		return query_list

#create a query function for data on multiple sheets

#create a query function that returns a cell or a list of data attributed to a specified unique identifier

def cell_return(file_name, ui_col_label, ui, ret_col_label):
	if type(ui) == list:
		return_list = []
		for ui in ui_list:
			return_list.append(cell_return(file_name, ui_col_label, ui, ret_col_label))
	else:
		return column_dict[file_name][ret_col_label][column_dict[file_name][ui_col_label].index(ui)]

#create a function that rewrites the value of a cell in a a column of a sheet based off of a list of unique identifiers for the row

def df_update(file_name, ui_list, update_col_label, update_value):
	for ui in ui_list:
		for label in list(column_dict[file_name].keys()):
			if ui in column_dict[file_name][label]:
				row_index = list(column_dict[file_name][label]).index(ui) - 1
				col_index = list(column_dict[file_name]).index(update_col_label)
				worksheet_dict[file_name].iloc[row_index][col_index] = update_value
	worksheet_dict[file_name].to_excel(r"Z:\Shared\Departments\BPI\Process Improvement\Records Retention\new-total-inventory-list.xlsx")

#create a function that returns a list of the cells unique only to the right column

def outlier(left_file_name, left_col_label, right_file_name, right_col_label):
	none_check_list = []
	none_check_list.append(column_dict[right_file_name][right_col_label][0]) #creates a list of the first cell value in specified column of right file
	outlier_list = [cell for cell in column_dict[right_file_name][right_col_label] if cell not in column_dict[left_file_name][left_col_label]]
	if (outlier_list == []) or (outlier_list == none_check_list):
		return print('None')
	else:
		return print(outlier_list)


		








