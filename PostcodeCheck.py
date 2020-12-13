#!/usr/bin/env python
import pandas as pd
import collections
print("running")
#CSV input file:
df = pd.read_csv("/Users/thomas/Desktop/Batch.csv")
#number of postcodes entered output
number_of_postcodes = int(len(df))
print ("Number of postcodes queried: ", number_of_postcodes)
print("Do NOT press any key...")
#change input csv to list
postcode_list= df.values.tolist()
#SIMD_2016_Postcode_lookup_tool return
path = "/Users/thomas/Documents/py/SIMDQuerytool/PostcodeCheck/2016Postcodes.xlsx"
df = pd.read_excel(path, sheet_name='All postcodes', usecols=[0])
postcodes=[]
postcode_exists=['Postcode exists',]
list_of_postcodes=[]
postcode_does_not_exist=['Postcode does NOT exist',]
requested_info_list=[]
#flatten list of list
for sublist in postcode_list:
    for item in sublist:
        list_of_postcodes.append(item)
#check if postcode exists
for postcode in list_of_postcodes:
    if postcode in df.values:
     postcode_exists.append(postcode)
    else:
        postcode_does_not_exist.append(postcode)
#extends list for output
lpostcode_exists,lpostcode_does_not_exist=len(postcode_exists), len(postcode_does_not_exist)
max_len= max(lpostcode_exists,lpostcode_does_not_exist)
if not max_len ==lpostcode_does_not_exist:
    postcode_does_not_exist.extend(['']*(max_len-lpostcode_does_not_exist))
#creates output list
requested_info_list.append(postcode_exists)
requested_info_list.append(postcode_does_not_exist)
#display and publish data
df1 = pd.DataFrame(requested_info_list[1:], columns=(requested_info_list[0]))
df2= df1.transpose()
df2.to_excel(r"/Users/thomas/Desktop/Requested_info_for_inputed_postcodes.xlsx", sheet_name='Requested Data')
