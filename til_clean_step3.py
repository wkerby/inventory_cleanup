import pandas as pd

#read the table of current box inventory from VRC into a df
vrc_box_inventory_as_of_01242022 = pd.read_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\vrc-box-inventory-as-of-01242022.xlsx')

#convert this box inventory df into a dictionary
vrc_box_inventory_as_of_01242022_dict = vrc_box_inventory_as_of_01242022.to_dict()

#pass a list of the range of the total number of rows in the vrc_box_inventory_as_of_01242022 df into a variable for ease of use
vrc_box_dict_range = list(range(len(vrc_box_inventory_as_of_01242022_dict[list(vrc_box_inventory_as_of_01242022_dict.keys())[0]].keys())))

#read the table of destroyed boxes as of 01-01-2017 into a df
vrc_destroyed_as_of_01012017 = pd.read_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\vrc-destroyed-as-of-01012017.xlsx')

#convert this destroyed box df into a dictionary
vrc_destroyed_as_of_01012017_dict = vrc_destroyed_as_of_01012017.to_dict()

#pass a list of the range of the total number of rows in the vrc_destroyed_as_of_01012017 df into a variable for ease of use
vrc_dest_dict_range = list(range(len(vrc_destroyed_as_of_01012017_dict[list(vrc_destroyed_as_of_01012017_dict.keys())[0]].keys())))

#add the actual date destroyed column to the vrc_box_inventory_as_of_01242022
#make sure to account for fact that there will be no 'DATE_DESTROYED' value for any rows that are currently in the vrc_box_inventory_as_of_01242022 because this is our current inventory
vrc_box_inventory_as_of_01242022_dict['DATE_DESTROYED'] = {}
for i in list(range(len(vrc_box_dict_range))):
	vrc_box_inventory_as_of_01242022_dict['DATE_DESTROYED'][i] = "" 

#take every row in the vrc_destroyed_as_of_01012017 table and add to the vrc_box_inventory_as_of_01242022
marker_start = len(vrc_box_dict_range)
column_add_dict = {'VRC_BARCODE':'Barcode Number', 'SBG':'fdivcode', 'DEPT_CODE':'fcust_dept','DESCRIPTION':'fbox_desc','DATE_DESTROYED':'feventtime'}
for dict_key in list(vrc_box_inventory_as_of_01242022_dict.keys()):
	for i in vrc_dest_dict_range:
			if dict_key == 'DATE_DESTROYED':
				vrc_box_inventory_as_of_01242022_dict[dict_key][marker_start+i] = str(vrc_destroyed_as_of_01012017[column_add_dict[dict_key]][i])[:-8]
				continue
			if dict_key in list(column_add_dict.keys()):
				vrc_box_inventory_as_of_01242022_dict[dict_key][marker_start+i] = vrc_destroyed_as_of_01012017_dict[column_add_dict[dict_key]][i]
			else:
				vrc_box_inventory_as_of_01242022_dict[dict_key][marker_start+i] = ""

#again pass a list of the range of the total number of rows in the vrc_box_inventory_as_of_01242022 df into a variable for ease of use
vrc_box_dict_range = list(range(len(vrc_box_inventory_as_of_01242022_dict[list(vrc_box_inventory_as_of_01242022_dict.keys())[0]].keys())))

#clean up the dates listed in the 'DATE_DESTROYED' column and change from YYYY-MM-DD format to MM/DD/YYYY format
century_list = ['20','19']
for i in vrc_box_dict_range:
	if str(vrc_box_inventory_as_of_01242022_dict['DATE_DESTROYED'][i])[:2] != '':
		if str(vrc_box_inventory_as_of_01242022_dict['DATE_DESTROYED'][i])[:2] in century_list:
			str_date_list = str(vrc_box_inventory_as_of_01242022_dict['DATE_DESTROYED'][i]).split("-")
			new_str_date = str_date_list[1] + '/' + str_date_list[2][:-1] + '/' + str_date_list[0]
			vrc_box_inventory_as_of_01242022_dict['DATE_DESTROYED'][i] = new_str_date
	else:
		pass

#retrieve all other columns of relevant data from Total Inventory List - Main (version 4).xlsx
#start by reading Total Inventory List - Main (version 4) into a df
total_inventory_list_main_version4 = pd.read_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main (version 4).xlsx')

#convert Total Inventory List - Main (version 4).xlsx df into a dictionary
total_inventory_list_main_version4_dict = total_inventory_list_main_version4.to_dict()

#now retrieve this data from Total Inventory List - Main (version 4) dictionary
#start by creating a list of columns from the Total Inventory List - Main (version 4) table that you would like to add to the vrc_box_inventory_as_of_01242022_dict
column_add_list = ['LOCATION','SBG','IsLegalHold','BUL_Comb_ID','BUL_1','BUL_2','BUL_3','Retention_Years','Status','Box_ID',
'IsDestroyed','IsPermanent','IsExtensionRequest','IsAsBuiltDrawings_','IsDrawings','IsStampedPlans','IsBlueprints','E1_Financial_Close_Date']

#initialize all values as blank for all added columns for each row-column intersection
for column in column_add_list:
	vrc_box_inventory_as_of_01242022_dict[column] = {}
for column in column_add_list:
	for i in vrc_box_dict_range:
		vrc_box_inventory_as_of_01242022_dict[column][i] = ""

#pass a list of the range of the total number of rows in the total_inventory_list_main_version4 df into a variable for ease of use
total_inv_dict_range = list(range(len(total_inventory_list_main_version4_dict[list(total_inventory_list_main_version4_dict.keys())[0]].keys())))

#add the corresponding values for each column onto the vrc_box_inventory_as_of_01242022_dict based off of vrc barcode
for i in vrc_box_dict_range:
	vrc_barcode = vrc_box_inventory_as_of_01242022_dict['VRC_BARCODE'][i]
	index = [row for row, barcode in list(total_inventory_list_main_version4_dict['VRC_Barcode'].items()) if str(barcode) == vrc_barcode]
	if len(index) > 0:
		for column in column_add_list[:-1]: #want to make sure not to include the 'Date_Destroyed' column, as it requires extra format
			vrc_box_inventory_as_of_01242022_dict[column][i] = total_inventory_list_main_version4_dict[column][index[0]]

#add the corresponding values for E1_Financial_Close_Date for each row using MM/DD/YYYY format
for i in vrc_box_dict_range:
	vrc_barcode = vrc_box_inventory_as_of_01242022_dict['VRC_BARCODE'][i]
	index = [row for row, barcode in list(total_inventory_list_main_version4_dict['VRC_Barcode'].items()) if str(barcode) == vrc_barcode]
	if len(index) > 0:
		date_str = str(total_inventory_list_main_version4_dict['E1_Financial_Close_Date'][index[0]]).split("-")
		if date_str != ['NaT']:
			date = date_str[1] + '/' + date_str[2][:2] + '/' + date_str[0]
			vrc_box_inventory_as_of_01242022_dict['E1_Financial_Close_Date'][i] = date
		else:
			pass

#make a df from the vrc_box_inventory_as_of_01242022_dict and write the table into an .xlsx file
final_df = pd.DataFrame(vrc_box_inventory_as_of_01242022_dict)
final_df.to_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main (version 5).xlsx', index = False)
