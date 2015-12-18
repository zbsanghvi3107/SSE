####################################################
#                  Revision: 1.1                   #
#              Updated on: 11/12/2015              #
#                                                  #
# What's new:                                      #
#                                                  #
# In Performance files, number of Disks assigned   #
# to each worker is dynamic.                       #
#                                                  #
####################################################


####################################################
#                  Revision: 1.0                   #
#              Updated on: 11/03/2015              #
####################################################

####################################################
#                                                  #
#   This script extracts IOps and Error data from  #
#   a specified directory and creates a .csv file. #
#                                                  #
#   Author: Zankar Sanghavi                        #
#                                                  #
#   © Dot Hill Systems Corporation                 #
#                                                  #
####################################################
import time


import os
# starting time, to calculate Elapsed Time.
start_time = time.time()

# to find path of Current Working Directory
# for Python scripts
c_path = os.getcwd()

import pandas
import csv

import numpy as np
import xlsxwriter

import sse_functions
import write_single_report

# accessing "sse_functions" fuction
# from "SSE_Functions.py" file
ssef= sse_functions.SSE_Functions

# accessing "write_single_report" function
# from "Write_Single_Report.py" file
wsr= write_single_report.Write_Single_Report


###################################
#  To get path of SSE and Baseline
#  files
###################################
path_1st = input('Please enter full path of a any SSE file: ')


#to check if it is a .csv file
while 'csv' not in path_1st:
    print('\nIt is not a .csv file!')
    path_1st = input('\nPlease enter full path of a any SSE file: ')

    
# find Directory from Performance files
file_dir_perf= os.path.dirname(r''+str(path_1st))
file_dir_perf= str(file_dir_perf).replace('"','') 

file_list_perf=os.listdir(file_dir_perf)


# Filter of none SSE files, so that we will have directory names 
dir_names =[file_list_perf[i] for i in range(len(file_list_perf))
            if 'SSE' not in file_list_perf[i] and '.txt' not in file_list_perf[i]]

           
# Filter of none SSE files, so that we will have directory paths 
dir_paths =[file_dir_perf+'\\'+file_list_perf[i] for i in range(len(file_list_perf)) 
           if 'SSE' not in file_list_perf[i] and '.txt' not in file_list_perf[i]]

           
# to alert user about number of Directories
print('\nFound ' +str(len(dir_paths))+ ' directories')


# to save report name for each directories
csv_name=[]
for d1 in range(len(dir_paths)):
    
    temp = input('\nSave Report for ' +str(dir_names[d1])+ ' as: ')
    csv_name.append(temp)


# to generate report for multiple directories
for d2 in range(len(dir_paths)): 

    file_list = os.listdir(dir_paths[d2])
    wsr.Generate_Single_Report(dir_paths[d2], file_list, file_dir_perf, file_list_perf, csv_name[d2])

    
# go back to original directory
os.chdir(r''+str(c_path)) 

# To show elapsed time   
elapse_time =round((time.time() - start_time),2) # in seconds
if elapse_time <= 60 :
    print("\nElapsed time: %s seconds" % elapse_time ) # in seconds
elif elapse_time > 60: # in minutes
    print("\nElapsed time: %s minutes" % round(((time.time() - start_time)/60),2))
    

#####################################
#              END                  #
#####################################    