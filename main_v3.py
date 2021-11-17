# a = 1.2
# b = 2.3

# print (a + b) 

print ('*********##########*********')

from posixpath import dirname
import pyodbc
from os import path

dir_path = path.dirname(path.realpath(__file__))
# print (dir_path)

from os import walk

from os import listdir
from os.path import isfile, join
onlyfiles = [f for f in listdir(dir_path) if (isfile(join(dir_path, f)))]
# print (onlyfiles)
from os import rename 
from shutil import copyfile
from time import sleep, perf_counter




####################################
#########Convert File EXT############
a = perf_counter()
for i in onlyfiles:
     if (i.find('.res') != -1):
        #  rename(i, (i.split('.res')[0] + '.accdb'))
        # print (i, " : ",  (i.split('.res')[0] + '.accdb'))
        copyfile(i, (i.split('.res')[0] + '.accdb'))

b = perf_counter()
c = b - a
print ('c1 : ', c) ##0.2415963

# ####################################
# #########Data collection############


a = perf_counter()
count=0
data_file_out = open("output_data.txt", "a")
onlyfiles = [f for f in listdir(dir_path) if (isfile(join(dir_path, f)))]
for i in onlyfiles:
    count = count + 1
#     print (i)
    if (i.find('.accdb') != -1):
        # print ("Contains given substring ")
        # print (i)
        
        db_file_path = dir_path + '\\' + i
        # print (db_file_path)
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+db_file_path)
        ## conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\tdnguy13\Documents\0.2 Batteries\Res_test\Exp1\accdb\acda.accdb')
        cursor = conn.cursor()

        ### Test_Name
        cursor.execute('select Test_Name from Global_Table')
        for row in cursor.fetchall():
            Test_Name = row

        ### Discharge_Capacity
        cursor.execute('select Discharge_Capacity from Channel_Normal_Table')
        for row in cursor.fetchall():
            Discharge_Capacity = row


        # print (Test_Name, "  ", str(Discharge_Capacity))    
        # print (type(str(Discharge_Capacity)))
        output = str(Test_Name) +  "  " + str(Discharge_Capacity) + '\n'
        # print (output)
        data_file_out.write(output)

        cursor.close()
        conn.close() 

#     else:
#         a = 0
#         # onlyfiles.remove(i)
#         # print ("Doesn't contains given substring")

data_file_out.write("***************** Print date" + " **************** \n")
data_file_out.close()

b = perf_counter()
c = b - a
print ('c : ', c)



#################################
#################################


# print (onlyfiles)

# db_file_path = dir_path + '\\' + 'acda.accdb'

# print ((db_file_path))

# conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+db_file_path)
# cursor = conn.cursor()

# ### Discharge_Capacity
# cursor.execute('select Discharge_Capacity from Channel_Normal_Table')
# for row in cursor.fetchall():
#     Discharge_Capacity = row
# # print ("Discharge_Capacity", Discharge_Capacity)


# ### Test_Name
# cursor.execute('select Test_Name from Global_Table')
# for row in cursor.fetchall():
#     Test_Name = row
# # print ("Test_Name", Test_Name)



##################################
#######Delete_extra_file############


from os import remove
a = perf_counter()
onlyfiles = [f for f in listdir(dir_path) if (isfile(join(dir_path, f)))]
for i in onlyfiles:
        # print (i)
        if (i.find('.accdb') != -1):
                # print (i)
                remove(i)
b = perf_counter()
c = b - a
print ('c1 : ', c)





from math import pi
print ('****************$$$$$$thuong$$$$$$$****************')
print (pi)
