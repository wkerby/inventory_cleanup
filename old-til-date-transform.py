# import csv
# file_name = r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\old-til-09-04-2019.csv'
# #read
# with open(file_name, "r", encoding = "utf-8-sig") as file_object: #add the utf-8-sig encoding argument to the function notebook
#     lines = list(csv.reader(file_object))
# for line in lines[1:]:
#         for i in list(range(0,len(line))):
#             string_list = list(line[i])
#             comma = ","
#             if comma in string_list:
#                 for z in list(range(0, len(string_list))):
#                     if string_list[z] == ",":
#                         string_list[z] = "."
#                 replace_string = ""
#                 for character in string_list:
#                     replace_string += character
#                     line.insert(i+1,replace_string)
#                     line.pop(i)
#             else:
#                 pass            
#         null_list = ["","#N/A","Yes"]
#         if line[19] in null_list:
#             line[19] = ""
#         else:
#             line[19] = line[19][:-5]

# #write
# with open(file_name, "w", encoding = "utf-8-sig") as file_object:
#    for line in lines:
#        for i in list(range(0,len(line))):
#            if i != len(line)-1:
#                file_object.write(line[i]+ ",")
#            else:
#                file_object.write(line[i] + "\n")
import pandas as pd
inner_join_list = []
right_anti_join_list = []
vrc_df = pd.read_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\vrc-destroyed-as-of-01012017.xlsx')
til_df = pd.read_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main (version 2).xlsm')
vrc_barcode_list = list(vrc_df.loc[:, 'Barcode Number'])
til_barcode_list = list(til_df.loc[:,"VRC_Barcode"])
for barcode in vrc_barcode_list:
    if barcode in til_barcode_list:
        inner_join_list.append(barcode)
    else:
        right_anti_join_list.append(barcode)
print("Percentage of matching barcodes: " + (str(len(inner_join_list)/(len(inner_join_list) + len(right_anti_join_list)))))

destruction_date_VRC_list = list(til_df.loc[:,"Destruction_Date_VRC"]) #list of destruction dates according to VRC
#print(destruction_date_VRC_list)
VRCcount = 0
for date in destruction_date_VRC_list:
    if date != pd.NaT:
        VRCcount += 1
date_destroyed_list = list(til_df.loc[:,"Date_Destroyed"]) #list of destruction dates according to B&G
BGcount = 0
for date in date_destroyed_list:
    if date != pd.NaT: 
        BGcount += 1
print("VRC Destructions: " + str(VRCcount) + "\nB&G Destructions: " + str(BGcount))
