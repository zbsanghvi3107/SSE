####################################################
#                  Revision: 1.0                   #
#              Updated on: 11/06/2015              #
####################################################

####################################################
#                                                  #
#   This script runs Batch file for IOMeter test   #
#   and renames the "result.csv" file to include   #
#   Serial Number and Model number.                #
#   Example: SERIALNUMBER_MODELNUMBER_sa_baseline  #
#                                                  #
#   Author: Zankar Sanghavi                        #
#                                                  #
#   © Dot Hill Systems Corporation                 #
#                                                  #
####################################################

import os
dir_path = os.getcwd()

# run batch from 
import subprocess
subprocess.Popen(r'"'+dir_path+'\SA_baseline_AF_KF.bat"')


# Wait for 12 minutes 30 seconds
time.sleep(750) # 12.5 * 60  seconds

##################
# Rename file
##################

#import shutil 
#import io

sn_path= r'' +dir_path+ '\sn.txt'
result_path = r''+dir_path + '\\result.csv'


################
# Reading sn.txt
################
with open(sn_path, encoding = 'utf-16-le') as sn:
    sn = sn.readlines()
    
# searching SerialNumber index in 1st line
line1 = sn[0]
sn_index = line1.find('SerialNumber', 0, len(line1))

# using sn_index to extract Serial Number from last line
ll_temp = sn[-1] #last line
sn = ll_temp[sn_index:] # serial number


# to remove extra spaces, commas and new lines(\n)
import re

sn = sn.strip('\n')
sn = re.sub( '\s+', ' ', sn ).strip()
sn = sn.replace(" ",',')

# extract data (model no.) that is situated between 1st and 2nd space.
model_no = ''

for i in range(len(ll_temp)):
    if ll_temp[i] == ' ':
        temp = i
        break
temp_model_no = ll_temp[temp+1:]


for j in range(len(temp_model_no)):
    if temp_model_no[j] == ' ':
        break
    model_no += str(temp_model_no[j])

# renaming result file to "SERIALNUMBER_MODELNUMBER_sa_baseline.csv"
os.rename(result_path, r''+dir + '\\' +sn+ '_' +model_no+ '_sa_baseline.csv')