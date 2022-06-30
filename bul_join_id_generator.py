import pandas as pd
from itertools import combinations
#this script will serve as a join table generator for any many-to-many relationship
#create the dictionary of combination id's and corresponding business unit leader combinations
max_bul_num = 4 #assume no more than 4 businnes unit leaders assigned to any one box or bag
bul_list = ['Bill Steed','Susan Stabler','Brian Murray','Joshua Hayes','Steve Manown','Chris Evans','Mike Foushee','Justin Rannick','Tate McKee','Ben Harris','Michael Thomas','Chip Grizzle','Wallace Watanabe','Bill Vaughn','Jeff Krogsgard','Hal Matern','Patti Kizziah','Ben Norton','Michael Stainback','Brig Eastman','Eric Perkinson','Ed Hauser','Tim Johnson','Reed Weigle','Stephen Franklin','Ben Barfield','Troy Ogden','Tom Garrett','Robby Hayes','Jason Hard','Katie Lewis','Randy Freeman','Matt Smith','Erik Sharpe','Jason Ellington','Jason Weeks','Wes Kelley','Chris Britton','Jill Deer','Brad Runyan','Drew Kelley','Kevin White','Michael Byrd','Trey Clegg','Jud Jacobs','Marty Hardin','Jim Ellspermann','Michael Keller','David Gordon','Peyton Robertson']
#bul_list = ['A','B','C']
bul_comb_list = []
id_bul_comb_dict = {}
for i in list(range(1,max_bul_num + 1)):
    for combination in list(combinations(bul_list,i)):
        bul_comb_string = ""
        for i in combination:
            if i == combination[-1]:
                bul_comb_string += i
            else:
                bul_comb_string += i + "-"
        bul_comb_list.append(bul_comb_string)
for i in list(range(1,len(bul_comb_list) + 1)):
    id_bul_comb_dict[i] = (bul_comb_list[i-1])
filename = 'bul_comb_join_table.csv'
with open(filename, 'w') as file_object:
    file_object.write("BUL_Comb_ID, BUL Combination\n")
    for ID, Comb in id_bul_comb_dict.items():
        #print(str(ID) + ': ' + Comb)
        file_object.write(str(ID) + ", " + str(Comb) + "\n")

#assign a bul combination id to boxes (either destroyed = Yes or destroyed = No) for which there is not yet an assigned BUL(s)
til = pd.read_excel(r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\Total Inventory List - Main (version 2).xlsm')
index = til.index
for barcode in list(til.loc[:,"VRC_Barcode"]):
    conditions = til["VRC_Barcode"] == barcode & til["BUL_1"].notnull() == True
    barcode_index = index[conditions]
    if til.loc[barcode_index,"BUL_2"].isnull() == True and til.loc[barcode_index,"BUL_3"].isnull() == True #only BUL_1
    if til.loc[barcode_index,"BUL_2"].isnull() == False and til.loc[barcode_index,"BUL_3"].isnull() == True #BUL_1 and BUL_2
    if til.loc[barcode_index,"BUL_2"].isnull() == False and til.loc[barcode_index,"BUL_3"].isnull() == False #<BUL_1 and BUL_2 and BUL_3
    bul_comb_test = list(til.loc[barcode_index, "BUL_1"])[0]
