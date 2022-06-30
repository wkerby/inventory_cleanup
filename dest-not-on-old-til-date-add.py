import csv
file_name = r'Z:\Shared\Departments\BPI\Process Improvement\Records Retention\dest-not-on-old-til.csv'
#read
with open(file_name, "r", encoding = "utf-8-sig") as file_object: #add the utf-8-sig encoding argument to the function notebook
    lines = list(csv.reader(file_object))
for line in lines[1:]:

        for i in list(range(0,len(line))):
            string_list = list(line[i])
            comma = ","
            if comma in string_list:
                for z in list(range(0, len(string_list))):
                    if string_list[z] == ",":
                        string_list[z] = "."
                replace_string = ""
                for character in string_list:
                    replace_string += character
                    line.insert(i+1,replace_string)
                    line.pop(i)
            else:
                pass            
        null_list = ["","#N/A","Yes"]
        if line[30] in null_list:
            pass
        else:
            if line[30][5] == str(0):
                line[19] = line[30][6] + '/' + line[30][8:10] + '/' +line[30][0:4]
            else:
                line[19] = line[30][5:7] + '/' + line[30][8:10] + '/' +line[30][0:4]
for line in lines:
    line[13] = "VRC"
    del line[27:33]
            

#write
with open(file_name, "w", encoding = "utf-8-sig") as file_object:
   for line in lines:
       for i in list(range(0,len(line))):
           if i != len(line)-1:
               file_object.write(line[i]+ ",")
           else:
               file_object.write(line[i] + "\n")