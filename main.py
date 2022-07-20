# -------------------------------------------------------------------------
# Made with LOVE By Gilad Gershon
# giladgershon.net
# --------------------------------------------------------------------------

import pandas as pd
import itertools 
import os
from tqdm import tqdm

# folder path
path = r'files'

#files list
file_list = []
count = 0


#count duplicates
dupcount = 0


for file in os.listdir(path):
    count +=1
    if os.path.isfile(os.path.join(path, file)):
     file_list.append(file)
     
     
    columns=['first_name','last_name','email','phone_numner'] #columns in the new xslx file
    first_name   = []                                         #First name column (new file)
    last_name    = []                                         #Last name column (new file)
    email        = []                                         #Email column (new file)
    phone_number = []                                         #Phone column (new file)
  
    count = 0

print('Starting to process '+ str(len(file_list)) +' files.')
for file in tqdm(file_list):    #for file in files folder
    file_name = path+'/'+file
    first_name_list = pd.read_excel(file_name  ,index_col=None, na_values=['NA'], usecols="A:A").values.tolist() #Collect first name
    last_name_list  = pd.read_excel(file_name  ,index_col=None, na_values=['NA'], usecols="B:B").values.tolist() #Collect last name
    email_list      = pd.read_excel(file_name  ,index_col=None, na_values=['NA'], usecols="D:D").values.tolist() #Collect first email
    phone_list      = pd.read_excel(file_name  ,index_col=None, na_values=['NA'], usecols="C:C").values.tolist() #Collect phone number

    for (fn,ln,e,p) in itertools.zip_longest(first_name_list, last_name_list, email_list, phone_list):  #for firstname, lastname, email and phone number merage unique rows to new xslx file
            
             
      if fn[0] not in first_name or ln[0] not in last_name or p[0] not in phone_number or e[0] not in email: #Collect only unique rows
        count +=1
        first_name.append(fn[0])
        last_name.append(ln[0])
        email.append(e[0])
        phone_number.append(str(p[0]))
      else:
        dupcount +=1 #count duplicates

              
print('Total unique rows that merge to the new file: ' + str(count))
print('We found ' + str(dupcount) + ' duplicates rows, we didnt copy them to the new file.')

#create new xslx file with the mergs rows
df = pd.DataFrame(list(zip(first_name,last_name,email,phone_number)), columns=columns) 
df.to_excel('new_file.xlsx', sheet_name='main')
print('Done')
