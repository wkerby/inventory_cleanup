#import pandas
import pandas as pd

#load total inventory list - main into a df
total_inv_df = pd.read_excel(r'Z:\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main.xlsm')

#create a list of all bul names that appear in BUL_1 column of Total Inventory List - Main
bul_names = []
for bul_name in list(total_inv_df.loc[:, "BUL_1"]):
	if type(bul_name) != float:
		if bul_name not in bul_names:
			bul_names.append(bul_name)
print(bul_names)
bul_names_ = {"BUL_Name":bul_names}
bul_names_ = pd.DataFrame(bul_names_)

bul_names_.to_excel(r'Z:\Departments\BPI\Process Improvement\Records Retention\bul-name-table.xlsx', index = False)

