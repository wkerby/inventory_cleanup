#create a script that generates a join table for barcodes and business unit leaders
#import the pandas library
import pandas as pd

#read the latest version of the total inventory list into a df
total_inventory_list_version5_df = pd.read_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main (version 5).xlsx')

#obtain a dictionary with barcodes as the keys and a list of its attributed
#business unit leader(s) as the corresponding values
#begin by obtaining a df that includes only current box inventory
index = total_inventory_list_version5_df.index
condition = total_inventory_list_version5_df["IsDestroyed"] == "No"
current_indices = index[condition]
current_df = total_inventory_list_version5_df.loc[current_indices, :]

#now obtain a list of the vrc barcodes for all current box inventory
current_inventory_barcodes = list(current_df.loc[:, 'VRC_BARCODE'])

#loop through the barcodes and check for a BUL_1, BUL_2, and BUL_3 attributed to each
# and add to a join dictionary
bul_barcode_join_dict = {}
for barcode in current_inventory_barcodes:
	bul_barcode_join_dict[barcode] = []

for barcode in current_inventory_barcodes:
	index = current_df.index
	condition = current_df['VRC_BARCODE'] == barcode
	barcode_index = index[condition]
	bul_bunch = ["BUL_1","BUL_2","BUL_3"]
	for bul in bul_bunch:
		bul_name = list(current_df.loc[barcode_index, bul])[0]
		if type(bul_name) != float:
			bul_barcode_join_dict[barcode].append(bul_name)
		else:
			pass

#take the bul_barcode_join_dict and make a new dictionary that will become a df
#i will then write this df to a .xlsx file, and it will serve as the join table
final_dict = {'VRC_BARCODE':{}, 'BUL_Name':{}}

#find the number of total indices that will appear as part of the join table by getting a 
#count value of the lengths of each bul_list for each barcode
range_num = 0
for barcode in list(bul_barcode_join_dict.keys()):
	range_num += len(bul_barcode_join_dict[barcode])

#initialize all key-value pairs in the final_dict dictionary
for i in list(range(range_num)):
	for column in list(final_dict.keys()):
		final_dict[column][i] = ''

#append barcode and bul_name combinations to final_dict dictionary
#begin first with barcodes
start_index = 0
for barcode, bul_name_list in list(bul_barcode_join_dict.items()):
	if len(bul_name_list) > 1:
		end_index = start_index + len(bul_name_list) -1
		for index in list(range(start_index,end_index+1)):
			final_dict['VRC_BARCODE'][index] = barcode
	else:
		end_index = start_index
		final_dict['VRC_BARCODE'][end_index] = barcode
	start_index = end_index + 1

#now append bul_names
for barcode_ in current_inventory_barcodes:
	#problem is that it is looping through all barcodes until it finds the 
	add_index_list = [index for index, barcode in list(final_dict["VRC_BARCODE"].items()) if barcode == barcode_]
	add_name_list = bul_barcode_join_dict[barcode_]
	# print(add_index_list)
	# print(add_name_list)
	if len(add_name_list) > 0:
		for index in add_index_list:
			final_dict["BUL_Name"][index] = add_name_list[add_index_list.index(index)]
	else:
		final_dict["BUL_Name"][add_index_list[0]] = "empty"

join_table = pd.DataFrame(final_dict)
join_table.to_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\vrc- barode-bul-join.xlsx')




