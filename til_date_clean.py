import pandas as pd
til = pd.read_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main (version 2).xlsm')

# #clean the dates that appear in the all date columns of the TIL so that their format is M/D/YYYY for single-digit months and single-digit days

#begin with the "Last_Acess_Date column"
til_barcodes = list(til.loc[:,'VRC_Barcode'])
index = til.index
for barcode in til_barcodes:
	iter_count = 0
	condition = til["VRC_Barcode"] == barcode
	barcode_index = index[condition] #this will retrieve a list of all indices in the dataframe that meet the specified condition
	date_str = list(str(list(til.loc[barcode_index, "Last_Access_Date"])[0]))
	while iter_count <= 1:
		if len(date_str) == 10 - iter_count or len(date_str) == 10:
			if iter_count == 0:
				if date_str[0] == str(0):			
					date_str.pop(0)
			else:
#clean the leading zero of the MM month and DD day section of the date
				if date_str[2] == str(0):
					#print(list(til.loc[barcode_index,"Last_Access_Date"])[0][3] + ' row ' + str(barcode_index[0]))
					date_str.pop(2)
		iter_count += 1
#erase the leading zero of the day section of the date
	if iter_count == 2:
		if len(date_str) == 10:
			if date_str[3] == str(0):
				date_str.pop(3)
		else:
			pass
		iter_count += 1
#clean up the 00:00:00 at the end of the dates in all date columns
	if iter_count == 3:
		if ":" in date_str:
			for i in list(range(9)):
				date_str.pop(-1)
		iter_count += 1
	new_character = ''
	for character in date_str:
		new_character += character
#re-arrange the order of month, date, and year for any incorrectly displayed dates in all date columns
	til.loc[barcode_index, "Last_Access_Date"] = new_character
	if iter_count == 4:
		date_str_test = ''
		for character in date_str:
			date_str_test += character
		if date_str_test[:4] in [str(i) for i in list(range(1950,2041))]:
			split_list = date_str_test.split('-')
			if split_list[1][0] == str(0) and split_list[2][0] == str(0):
				date_str_test = split_list[1][1] + '/' + split_list[2][1] + '/' + split_list[0]
			elif split_list[1][0] != str(0) and split_list[2][0] == str(0):
				date_str_test = split_list[1] + '/' + split_list[2][1] + '/' + split_list[0]
			elif split_list [1][0] != str(0) and split_list[2][0] != str(0):
				date_str_test = split_list[1] + '/' + split_list[2] + '/' + split_list[0]
			else:
				date_str_test = split_list[1][1] + '/' + split_list[2] + '/' + split_list[0]
	til.loc[barcode_index, "Last_Access_Date"] = date_str_test
#save changes to the data in a new .xlsx file
til.to_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main (version 3).xlsx')

#make a new df representing data in version 3 of the total inventory list
til = pd.read_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main (version 3).xlsx')

#re-arrange format of "Date_Destroyed" column in total iventory list
for barcode in til_barcodes:
	condition = til["VRC_Barcode"] == barcode
	barcode_index = index[condition]
	for column in ["Date_Destroyed","E1_Financial_Close_Date","Destruction_Date_VRC"]: 
		date_str = list(str(list(til.loc[barcode_index, column])[0]))
		delimit = ""
		if delimit.join(date_str)[:3] == 'NaT' or 'NaN':
			pass
		if delimit.join(date_str)[:4] in [str(i) for i in list(range(1950,2041))]: #basically stating if the value in the "Date_Destroyed" column begins with a YYYY value for this barcode
			for i in list(range(9)):
				date_str.pop(-1) #add the line of code from date_str_test
			new_character = ''
			for character in date_str:
				new_character += character
			split_list = new_character.split('-')
			if split_list[1][0] == str(0) and split_list[2][0] == str(0):
				new_character = split_list[1][1] + '/' + split_list[2][1] + '/' + split_list[0]
			elif split_list[1][0] != str(0) and split_list[2][0] == str(0):
				new_character = split_list[1] + '/' + split_list[2][1] + '/' + split_list[0]
			elif split_list [1][0] != str(0) and split_list[2][0] != str(0):
				new_character = split_list[1] + '/' + split_list[2] + '/' + split_list[0]
			else:
				new_character = split_list[1][1] + '/' + split_list[2] + '/' + split_list[0]
		til.loc[barcode_index, column] = new_character
til.to_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main (version 4).xlsx')




