import pandas as pd
til_df = pd.read_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main (version 4).xlsx')
#create a filtered view of data frame that includes only boxes that are not permanently kept, are not destroyed, and have a blank retention years column