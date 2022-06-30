import pandas as pd

#make a dictionary of all bul names and bul destruction report file paths
#use this to assign BULs to destroyed boxes for which there is not yet an attributed BUL
#use this with all destruction reports attributed to all BULs
#bul_anamole = ['Ben Barfield', 'Ben Harris', 'David Gordon', 'Ed Hauser','Jason Hard','Jason Weeks','Jill Deer','Jim Ellspermann','Randy Freeman','Trey Clegg','Troy Ogden', 'Wes Kelley']
bul_dest_report_dict = {}
bul_list = ['Ben Harris','Ben Norton','Bill Steed','Bill Vaughn','Brad Runyan','Brig Eastman','Chip Grizzle','Chris Britton','Chris Evans','Drew Kelley','Eric Perkinson','Erik Sharpe','Hal Matern','Jason Ellington','Jeff Krogsgard','Jud Jacobs','Justin Rannick','Katie Lewis','Katie Voss','Kevin White','Marty Hardin','Matt Smith','Michael Byrd','Michael Keller','Michael Stainback','Michael Thomas','Mike Foushee','Patti Kizziah','Reed Weigle','Robby Hayes','Stephen Franklin','Steve Manown','Susan Stabler','Tate McKee','Tim Johnson','Tom Garrett','Wallace Watanabe']
bul_dest_filepath_list = [r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Ben Harris.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Ben Norton.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Bill Steed.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Bill Vaughn.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Brad Runyan.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Brig Eastman.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Chip Grizzle.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Chris Britton.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Chris Evans.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Drew Kelley.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Eric Perkinson.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Erik Sharpe.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Hal Matern.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Jason Ellington.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Jeff Krogsgard.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Jud Jacobs.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Justin Rannick.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Katie Lewis.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Katie Voss.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Kevin White.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Marty Hardin.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Matt Smith.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Michael Byrd.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Michael Keller.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Michael Stainback.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Michael Thomas.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Mike Foushee.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Patti Kizziah.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Reed Weigle.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Robby Hayes.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Stephen Franklin.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Steve Manown.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Susan Stabler.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Tate McKee.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Tim Johnson.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Tom Garrett.xlsx',
r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Destruction Reports by BUL\Destruction Report - Wallace Watanabe.xlsx']
for i in list(range(0,len(bul_list))):
	bul_dest_report_dict[bul_list[i]] = bul_dest_filepath_list[i]

#read the total inventory list into python via pandas 
til = pd.read_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main (version 4).xlsx')

#define a function to retrieve a list of barcodes from a a destruction report
def retrieve_list(bul_dest_report_path, barcode_col_name):
	bul_destruction_df = pd.read_excel(bul_dest_report_path)
	return list(bul_destruction_df.loc[:, barcode_col_name])

#add barcodes to this dictionary if they are not currently in the total inventory list. this will help track down barcodes for boxes that were destroyed prior to the oldest version of the til
vrc_barcode_not_found = {bul_name: [] for bul_name in bul_list} 

#define a function that updates til with correct BUL if current BUL not equal to that listed on the BUL-specific destruction report
def verify_bul(bul_name, barcode_list):
	box_df = til.loc[til['VRC_Barcode'].isin(barcode_list)]
	for barcode in barcode_list:
		if len(list((box_df.loc[box_df['VRC_Barcode'] == barcode].loc[:,"BUL_1"]))) > 0: #if the BUL_1 Field is not blank for the row for this barcode. This is throwing an index error
			if list(box_df.loc[box_df['VRC_Barcode'] == barcode].loc[:,"BUL_1"])[0] != bul_name:
				index_val = list((til.loc[(til['VRC_Barcode'].isin([barcode]))].index))[0]
				til.loc[index_val, "BUL_1"] = bul_name
		else:
			# list(vrc_barcode_not_found.keys()).append(bul_name)
			vrc_barcode_not_found[bul_name].append(barcode)
			#print(str(barcode) + ' does not exist in the til')

for bul_name, bul_dest_filepath in list(bul_dest_report_dict.items()):
	verify_bul(bul_name, retrieve_list(bul_dest_filepath,'VRC_BARCODE'))

til.to_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main (version 5).xlsx')

			
# #retrieve a list of all barcodes for which there is a destruction date but no assigned BUL
# dest_blank_bul_df = til.loc[(til['BUL_1'].isnull()) & (til['IsDestroyed'] == 'Yes')]
# dest_blank_bul_barcodes = list(til.loc[(til['BUL_1'].isnull()) & (til['IsDestroyed'] == 'Yes')].loc[:,'VRC_Barcode'])





