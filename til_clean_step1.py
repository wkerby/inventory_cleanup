import pandas as pd

#read both worksheets as dfs  
vrc_df = pd.read_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\vrc-destroyed-as-of-01012017.xlsx')
til_df = pd.read_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main (version 2).xlsm') 

#retrieve a list of all VRC_Barcodes for boxes destroyed as of 01-01-2017
vrc_dest_barcode_list = list(vrc_df.loc[:,"Barcode Number"])

#retrieve a list of all VRC Barcodes for boxes that have been destroyed according to the TIL but that have no Destruction_Date
destroyed_til_df = til_df.loc[til_df["IsDestroyed"] == "Yes"] #retrieve all barcodes for boxes that have been destroyed
destroyed_no_date_til_df = destroyed_til_df.loc[destroyed_til_df["Date_Destroyed"].isnull()] #retrieve all barcodes for boxes that have been destroyed and for which there is no recorded "Destruction_Date" 
index_list = destroyed_no_date_til_df.index
til_dest_empty_dest_date_barcode_list = list(til_df.loc[index_list, 'VRC_Barcode'])

#retrieve a list that is the overlap (date_add_list) between the til_dest_empty_dest_date_barcode_list and the vrc_dest_barcode_list
#retrieve a list that represents all boxes that have been destroyed per vrc's destruction report but are not on the til (i.e. dest_not_on_til_list)
date_add_list = []
dest_not_on_til_list = []
for barcode in vrc_dest_barcode_list:
	if barcode in til_dest_empty_dest_date_barcode_list:
		date_add_list.append(barcode)
	else:
		pass
for barcode in vrc_dest_barcode_list:
	if barcode not in list(destroyed_til_df.loc[:,"VRC_Barcode"]):
		dest_not_on_til_list.append(barcode)

#read Total Inventory List - Main (version 2) as a csv
#assign the destruction date attributed to the VRC Barcode in vrc-destroyed-as-of-01012017.xlsx to the empty date attribute for the same VRC Barcode in Total Inventory List - Main (version 2).xlsm
for barcode in date_add_list:
	index = vrc_df.loc[vrc_df["Barcode Number"] == barcode].index
	date_fill = list(vrc_df.loc[index,"feventtime"])[0]

	index = til_df.loc[til_df['VRC_Barcode'] == barcode].index
	til_df.loc[index, 'Date_Destroyed'] = date_fill

#create new entries for VRC barcodes that were destroyed as of 01-01-2017 but are not listed on the til
til_header_list = [] #create a list of the column headers in the til df
for column in til_df.columns:
	til_header_list.append(column)

#create a dictionary that will store all data for vrc boxes that have been destroyed but are not yet on the til
for barcode in dest_not_on_til_list:
	new_entry = {} 
	for header in til_header_list: #loop through list of headers while creating the new_entries dictionary, because all but VRC_BARCODE and Date_Destroyed will be blank
		if header == "VRC_Barcode":
			new_entry[header] = barcode
		elif header == "Date_Destroyed": #assign the "feventime" from the vrc barcode destruction list attributed to the barcode to the Date_Destroyed field for the box's record
			index = vrc_df.loc[vrc_df["Barcode Number"] == barcode].index #find the index value of the barcode from right_anti_join_list within the vrc_df
			new_entry[header] = list(vrc_df.loc[index, "feventtime"])[0] #return the destruction date (feventtime) of this barcode within vrc_df
		elif header == "Destruction_Year":
			index = vrc_df.loc[vrc_df["Barcode Number"] == barcode].index
			new_entry[header] = str(list(vrc_df.loc[index, "feventtime"])[0])[:4]
		elif header == "SBG":
			index = vrc_df.loc[vrc_df["Barcode Number"] == barcode].index
			new_entry[header] = list(vrc_df.loc[index, "fdivcode"])[0]
		elif header == "IsLegalHold" or header == "IsPermanent":
			new_entry[header] = "No"
		elif header == "IsDestroyed":
			new_entry[header] = "Yes"
		elif header == "Description":
			index = vrc_df.loc[vrc_df["Barcode Number"] == barcode].index
			new_entry[header] = list(vrc_df.loc[index, "fbox_desc"])[0]
		else:
			new_entry[header] = ""

	til_df = til_df.append(new_entry, ignore_index = True) #append this new_entry dictionary as an entry in the til_df

#write the new data frame to excel as Total Inventory List - Main (version 3)
til_df.to_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main (version 4).xlsx', index = False)

# index = til_df.index
# for barcode in list(til_df.loc[index, "VRC_Barcode"])[:5]:
# 	index = til_df.loc[til_df['VRC_Barcode'] == barcode].index
# 	print(til_df.loc[index, "Last_Access_Date"])
	


