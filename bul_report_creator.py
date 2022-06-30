import pandas as pd

#try reading current inventory list into df and then into a dictionary and make user aware of error if otherwise
try:
	inv = pd.read_excel(r'Z:\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main_.xlsx')
	inv_dict = inv.to_dict()
except FileNotFoundError:
	print("Program doesn't recognize the file.")
else:
	print("File was recognized!")

#try reading bul join table into df and then into a dictionary and make user aware of error if otherwise
try:
	bul_join = pd.read_excel(r'Z:\Departments\BPI\Process Improvement\Records Retention\vrc-barode-bul-join.xlsx')
	bul_join_dict = bul_join.to_dict()
except FileNotFoundError:
	print("Program doesn't recognize the file.")
else:
	print("File was recognized!")

# #try reading an old til that includes job numbers into df and then into a dictionary and make user aware of error if otherwise
# try:
# 	old_inv = pd.read_excel(r'Z:\Departments\BPI\Process Improvement\Records Retention\vrc-barode-bul-join.xlsx')
# 	old_inv_dict = old_inv.to_dict()
# except FileNotFoundError:
# 	print("Program doesn't recognize the file.")
# else:
# 	print("File was recognized!")

#create a list of all bul names 
buls = ['Bill Steed', 'Susan Stabler', 'Steve Manown', 'Chris Evans', 'Mike Foushee', 'Justin Rannick', 'Tate McKee', 'Ben Harris', 'Michael Thomas', 'Chip Grizzle', 'Wallace Watanabe', 'Bill Vaughn', 'Jeff Krogsgard', 'Hal Matern', 
'Patti Kizziah', 'Ben Norton', 'Michael Stainback', 'Brig Eastman', 'Eric Perkinson', 'Ed Hauser', 'Tim Johnson', 'Reed Weigle', 'Stephen Franklin', 'Ben Barfield', 'Troy Ogden', 
'Tom Garrett', 'Robby Hayes', 'Jason Hard', 'Katie Lewis', 'Randy Freeman', 'Matt Smith', 'Erik Sharpe', 'Jason Ellington', 'Jason Weeks', 'Wes Kelley', 'Chris Britton', 'Jill Deer', 
'Brad Runyan', 'Drew Kelley', 'Kevin White', 'Michael Byrd', 'Trey Clegg', 'Jud Jacobs','Jim Ellspermann', 'Michael Keller', 'David Gordon', 'Katie Voss', 'Marty Hardin']

#sift out barcodes for which they are multiple buls
#this will be used in the last destruction report
repeat_barcodes = []
unique_barcodes = []
for vrc_barcode in list(bul_join_dict['VRC_BARCODE'].values()):
	if vrc_barcode not in unique_barcodes:
		unique_barcodes.append(vrc_barcode)
	else:
		repeat_barcodes.append(vrc_barcode)
for vrc_barcode in repeat_barcodes:
	if vrc_barcode in unique_barcodes:
		unique_barcodes.remove(vrc_barcode)

#create a dict of bul names and attributed unique barcodes (omit barcodes to which multiple buls are attributed)
bul_barcode_dict ={}
bul_filter_dict = {}
for bul in buls:
	bul_barcode_dict[bul] = []
	bul_filter_dict[bul] = []
for vrc_barcode in unique_barcodes:
	for index, barcode in list(bul_join_dict['VRC_BARCODE'].items()):
		if barcode == vrc_barcode:
			bul_barcode_dict[bul_join_dict['BUL_Name'][index]].append(vrc_barcode)

#loop through each list of barcodes in bul_barcode_dict
#check to see if it dest_date in inventory dict is in or before the year 2021
#filter out extension request, legal hold, and any other permanent boxes 
#do this for each bul until there are no longer any barcodes to be filtered
for bul in list(bul_barcode_dict.keys()):
	active = True
	iter_count = 1
	while active == True:
		start = len(bul_barcode_dict[bul])
		float_count = 0
		date_range = 0
		legal = 0
		extension = 0
		permanent = 0
		for vrc_barcode in bul_barcode_dict[bul]:
			index = [index for index, barcode in list(inv_dict["VRC_BARCODE"].items()) if barcode == vrc_barcode][0]
			if type(inv_dict['DESTRUCTION_DATE'][index]) == float:
				bul_filter_dict[bul].append(vrc_barcode)
				bul_barcode_dict[bul].remove(vrc_barcode)
				float_count += 1
			elif int(inv_dict['DESTRUCTION_DATE'][index][-4:]) > 2021:
				bul_filter_dict[bul].append(vrc_barcode)
				bul_barcode_dict[bul].remove(vrc_barcode)
				date_range += 1
			elif inv_dict["IsLegalHold"][index] == "Yes":
				bul_filter_dict[bul].append(vrc_barcode)
				bul_barcode_dict[bul].remove(vrc_barcode)
				legal += 1
			elif inv_dict["IsExtensionRequest"][index] == "Yes":
				bul_filter_dict[bul].append(vrc_barcode)
				bul_barcode_dict[bul].remove(vrc_barcode)
				extension += 1
			elif inv_dict["IsPermanent"][index] == "Yes":
				bul_filter_dict[bul].append(vrc_barcode)
				bul_barcode_dict[bul].remove(vrc_barcode)
				permanent += 1
			else:
				pass
		finish = len(bul_barcode_dict[bul])
		print(bul + " filter iteration " + str(iter_count) + ":")
		print("Float: " + str(float_count))
		print("Date: " + str(date_range))
		print("Legal: " + str(legal))
		print("Extension: " + str(extension))
		print("Permanent: " + str(permanent))
		print("Number filtered: " + str(start - finish))
		if start - finish == 0:
			active = False
		else:
			iter_count += 1

#create a dictionary for each bul that stores all barcodes that meet the filter criterion along with box description, department code, retention years, and bul
for bul in list(bul_barcode_dict.keys()):
	bul_write_dict = {"VRC_BARCODE":bul_barcode_dict[bul],"DESTRUCTION_DATE":[],"IsExtensionRequest":[],"IsLegalHold":[],"IsPermanent":[],"DESCRIPTION":[],"BUL_Name":[]}
	for i in list(range(len(bul_barcode_dict[bul]))):
		bul_write_dict["BUL_Name"].append(bul)
	for vrc_barcode in bul_barcode_dict[bul]:
		index = [index for index, barcode in list(inv_dict["VRC_BARCODE"].items()) if barcode == vrc_barcode][0]
		for column in list(bul_write_dict.keys())[1:6]:
			bul_write_dict[column].append(inv_dict[column][index])
	# print(bul_write_dict)
	bul_write_df = pd.DataFrame(bul_write_dict)
	try:
		filename = f'Z:\\Departments\\BPI\\Process Improvement\\Records Retention\\Emailed Records List\\2022\\{bul} Paper Records Due for Destruction.xlsx'
		bul_write_df.to_excel(filename,index = False)
		# bul_write_df.to_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Emailed Records List\2022\Susan Stabler Paper Records Due for Destruction.xlsx',index=False) 
	except FileNotFoundError:
		print("Program doesn't recognize the file!")
	else:
		print("Dictionary written to Excel.")








