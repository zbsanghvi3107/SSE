####################################################
#                  Revision: 1.2                   #
#              Updated on: 11/13/2015              #
#                                                  #
# What's new:                                      #
#                                                  #
# CONDITIONAL FORMATING:                           #
#                                                  #
#   Font color will change if it matches given     # 
#   conditions.                                    #
#                                                  #
#   Added two functions to support it:             #
#                                                  #
#       1.) add_conditional_formatting             #
#       2.) rank                                   #
#                                                  #
####################################################


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
#   This script contains all functions usefull     #
#   to generate a report for SSE Test              #
#                                                  #
#   Author: Zankar Sanghavi                        #
#                                                  #
#   © Dot Hill Systems Corporation                 #
#                                                  #
####################################################
import pandas
import numpy as np

class SSE_Functions:
    
    ########################################
    # This function will extract data which 
    # is only present in data_list from
    # sn_dic. 
    ########################################
    def extract_present_data(dictionary1, dictionary2, dictionary3, data_list):
        extracted_values1 = []
        extracted_values2 = []
        extracted_values3 = []
        extracted_keys = []
        #data_list = serial_nos
        for i in range(len(list(dictionary1.values()))):
            for j in range(len(data_list)):
                if str(list(dictionary1.values())[i]) == str(data_list[j]):
                    extracted_values1.append(list(dictionary1.values())[i])
                    extracted_values2.append(list(dictionary2.values())[i])
                    extracted_values3.append(list(dictionary3.values())[i])
                    extracted_keys.append(list(dictionary1.keys())[i])
                    #break
        return [extracted_values1, extracted_values2, extracted_values3, extracted_keys]
    
    
    
    ###################################
    # This function will find ".txt"
    # file from the File list and 
    # generate Serial Number and
    # Model Number associated with their
    # Physical Drive number.
    ###################################
    def generate_sn_mn_dictionary(file_dir, file_list):
        
        sn_name = [file_list[i] for i in range(len(file_list)) 
                                if "sn_multiple_drives.txt" in file_list[i]]
        
        sn_path = r'' +file_dir+ '\\' +sn_name[0]
        
        
        ################
        # Reading sn.txt
        ################
        with open(sn_path, encoding = 'utf-16-le') as sp:
            sp = sp.readlines()
            
        # searching SerialNumber
        line1 = sp[0]
        sn_index = line1.find('SerialNumber', 0, len(line1))
        
        # searching Name
        name_index = line1.find('Name', 0, len(line1))
        name_index
        
        import re
        
        vn_dic = {} # empty dictionary for Vendor Name
        sn_dic = {} # empty dictionary for Serial Number
        mn_dic = {} # empty dictionary for Model Number


        for i in range(2,len(sp)):
            line_temp = sp[i] 
            # access line from 2nd(if it starts 
            # from  0th ) line onwards
            
            ##############################    
            # to extract Vendor Name
            ##############################
            for j in range(len(line_temp)):
                if line_temp[j] == ' ':
                    temp = j
                    break
            temp_vendor = line_temp[:temp]
            
                
            ##############################    
            # to extract model number
            ##############################    
            model_no = ''

            for j in range(len(line_temp)):
                if line_temp[j] == ' ':
                    temp = j
                    break
            temp_model_no = line_temp[temp+1:]


            for j in range(len(temp_model_no)):
                if temp_model_no[j] == ' ':
                    break
                model_no += str(temp_model_no[j])
            #print(model_no)


            ##############################    
            # to extract serial number
            ##############################
            sn = line_temp[sn_index-1:] 

            # to strip extra characters
            sn = sn.strip('\n')
            sn = re.sub( '\s+', ' ', sn ).strip()
            sn = sn.replace(" ",',')
            #print(sn)

            ##############################    
            # to extract serial number
            ##############################
            phy_drive = line_temp[name_index:] # physical drive number

            for j in range(len(phy_drive)):
                if phy_drive[j] == ' ':
                    temp = j
                    break
            temp_phy_drive = phy_drive[:temp]
            #print(temp_phy_drive)

            # extracting only drive numbers
            temp_phy_drive = re.findall(r'\d+', temp_phy_drive)
            #print(temp_phy_drive)

            # To generate dictonary for  Vendor Name
            temp_vn = {int(temp_phy_drive[0]) : str(temp_vendor)}
            vn_dic.update(temp_vn)
                    
            # To generate dictonary for  Serial Number
            temp_dic_sn = {int(temp_phy_drive[0]) : str(sn)}
            sn_dic.update(temp_dic_sn)

            # To generate dictonary for  Serial Number
            temp_dic_mn = {int(temp_phy_drive[0]) : str(model_no)}
            mn_dic.update(temp_dic_mn)
        #print(sn_dic)    
        return [vn_dic, sn_dic, mn_dic]
    
    
    
    ######################
    # This function will
    # extract IOps &
    # Error data from
    # all the "baseline"
    # files.
    ######################
    def get_data_baseline(file_list, file_dir, sn_dic, mn_dic):
        from decimal import Decimal, ROUND_HALF_UP
        
        iops_final_list=[]
        errors_final_list= []
        
        import os
        import sys
        
        ###################################
        #  Importing from other Directory
        ###################################
        os.chdir('..')
        c_path = os.getcwd()
        sys.path.insert(0, r''+str(c_path)+'/Common Scripts')
        
        import report_functions
        rf= report_functions.Report_Functions

        sys.path.insert(0, r''+str(c_path)+'/SSE')
        
        
        #################################
        # Filter of 'baseline', 
        # just to read baseline .csv 
        # files
        #################################
        file_list =[file_list[i] for i in range(len(file_list)) if 
                   'baseline' in file_list[i]]
                   
        
        ########################
        # Finding serial nos to
        # cross check with 
        # Physical Drives
        ########################
        serial_nos=[]

        for i in range(len(file_list)): 
            file_name = file_list[i]
            for j in range(len(file_name)):
                if file_name[j] == '_':
                    temp_index = j
                    break
            serial_nos.append((file_name[:temp_index]))
        
   
        ########################
        # This will sort file 
        # list as per the 
        # drive number associated
        # with Serial Number.
        ########################
        sorted_files=[]

        for i in range(len(list(sn_dic.values()))):
            for j in range(len(file_list)):
                if list(sn_dic.values())[i] in file_list[j]:
                    sorted_files.append(file_list[j])
                    break
                 
        #print(sorted_files)
        
        
        #################################
        # Extractting 1st Time Stamp 
        # info
        #################################
        csv_data1= pandas.read_csv( r''+str(file_dir)+'\\'
                   + str(sorted_files[0])
                   , nrows=6
                   , header= None )
                          
        time_stamp_1st= csv_data1[0][5]
        
        
        #################################
        # Extractting 2nd Time Stamp 
        # info
        #################################
        csv_data2= pandas.read_csv( r''+str(file_dir)+'\\'
                                   + str(sorted_files[0])
                                   , skiprows=13
                                   , header= None )
        
        [ts_count, ts_indices]= rf.find_string(csv_data2
                                                , 0, 0
                                                ,'\'Time Stamp')
        
        time_stamp_2nd= csv_data2[0][ts_indices[0]+1]
        
        final_time_stamps=np.hstack((time_stamp_1st
                            ,time_stamp_2nd))
        final_time_stamps=final_time_stamps.tolist()

        final_time_stamps=['Time Stamp']+final_time_stamps
        
        
        #####################
        # Extract Test names 
        # for different test
        # Say: 4k & 512k
        # from the 1st file.
        # 
        # It is assumed that
        # Test Description
        # remains same
        # throughout all
        # baseline and
        # performance files.
        #####################
        j=0
        [worker_count, worker_indices]= rf.find_string(csv_data2,0,0,'WORKER')
        
        test_name_list= [csv_data2[2][worker_indices[j]] 
                      for j in range(len(worker_indices)) ] 
                      
        test_name_list=['Test Description']+test_name_list
        
        
        no_of_files= len(sorted_files)
        
        # Initializing 
        total_iops_1st=0
        total_iops_2nd=0
        total_errors_1st=0
        total_errors_2nd=0
        
        for k in range(no_of_files):
        
            csv_data2= pandas.read_csv( r''+str(file_dir)+'\\'
                           + str(sorted_files[k])
                           , skiprows=13
                           , header= None )
            
            [worker_count, worker_indices]= rf.find_string(csv_data2,0,0,'WORKER')

            
            ###################
            # IOps value for 4k
            # & 512k from 7th 
            # column.  
            #
            # It is assumed 
            # it there will be 
            # IOps data in 7th 
            # column for all 
            # files.
            ###################
            iops_list= [ round(float(csv_data2[6][worker_indices[j]]),2)
                        for j in range(len(worker_indices)) ] 
                        
            #1st element i.e. value of 4k
            total_iops_1st+= float(iops_list[0]) 
            
            #2nd element i.e. value of 512k
            total_iops_2nd+= float(iops_list[1])  
            
            #appending Drive# numbers for Report
            for i in range(len(list(sn_dic.values()))):
                
                if list(sn_dic.values())[i] in sorted_files[k]: 
                    temp_drive = r'Drive#' +str(list(sn_dic.keys())[i])
                    break
                    
            iops_list=[r''+str(temp_drive)] + iops_list
            
            
            ######################
            # Errors value for 4k
            # & 512k from 24th 
            # column
            ######################
            errors_list= [int(float(csv_data2[23][worker_indices[j]])) 
                          for j in range(len(worker_indices)) ]
            
            #1st element i.e. value of 4k
            total_errors_1st+= float(errors_list[0]) 
            
            #2nd element i.e. value of 512k
            total_errors_2nd+= float(errors_list[1]) 
                      
            # appending Drive# numbers for Report          
            errors_list=[r''+str(temp_drive)]+errors_list              
                                      
            # converting into an array for Better control
            # over data.
            iops_list= np.array(iops_list)
            errors_list= np.array(errors_list)

            iops_final_list.append(iops_list)
            errors_final_list.append(errors_list)
            
        #1st test, IOps average #i.e. value of 4k
        avg_iops_1st= round((total_iops_1st/no_of_files),2)
                
        #2nd test, IOps average #i.e. value of 512k
        avg_iops_2nd= round((total_iops_2nd/no_of_files), 2)  
        
        
        #1st test, Errors average
        avg_errors_1st= int(float(total_errors_1st/no_of_files))
        
        #2nd test, Errors average
        avg_errors_2nd= int(float(total_errors_2nd/no_of_files))
        
        
        # appending System text for Report, IOps  
        system_iops_list= ['System '+str(avg_iops_1st)+' '+str(avg_iops_2nd)]
        system_iops_list= np.array(system_iops_list)
        iops_final_list= [system_iops_list]+iops_final_list

        
        # appending System text for Report, Errors        
        system_errors_list=['System '+str(avg_errors_1st)+' '+str(avg_errors_2nd)]
        system_errors_list= np.array(system_errors_list)
        errors_final_list= [system_errors_list]+errors_final_list

        return [test_name_list, final_time_stamps, iops_final_list, errors_final_list, no_of_files, serial_nos]
        
        
        
    #################################
    # This function get iops and 
    # errors, test names and 
    # time_stamp information and
    # write a .csv file in where
    # all "baseline" csv file are
    # situated.
    #################################
    def create_csv_sse(csv_name, file_dir, test_name_list,
                      final_time_stamps, iops_final_list
                      , errors_final_list):
                      
                      
        #################################
        #   Header 
        #################################
        ioresult_str= 'IOps Results,,,'
        iometer_str= 'IOMeter Errors,,,'

        with open(r'' +str(file_dir)+'\\' + str(csv_name)+'.csv','w') as out_file:

            out_string= ''
            out_string+= str(ioresult_str)  +'\n'
            
            # Replacing extra characters, spaces and symbols to match
            # csv format
            temp= (str(final_time_stamps)).strip('[')
            temp= temp.strip(']')
            temp= temp.strip('\'')
            temp= temp.replace("'",'')
           
            temp1= (str(test_name_list)).strip('[')
            temp1= temp1.strip(']')
            temp1= temp1.strip('\'')
            temp1= temp1.replace("'",'')
                        
            out_string+= '\n' + temp
            out_string+= '\n' + temp1
            
            
            for i in range(len(iops_final_list)):
                temp= str(iops_final_list[i])
                temp= temp.strip('[')
                temp= temp.strip(']')
                temp= temp.strip('\'')
                temp= temp.replace("'",'')
                
                temp= temp.replace(" ",',')
                out_string+= '\n' + temp 
                
            out_string+= '\n\n' + str(iometer_str) +'\n'
            
            for j in range(len(errors_final_list)):
            
                temp=str(errors_final_list[j])
                temp= temp.strip('[')
                temp= temp.strip(']')
                temp= temp.strip('\'')
                temp= temp.replace("'",'')

                temp= temp.replace(" ",',')
                out_string+= '\n' + temp
                
            out_file.write(out_string)
    

    
    ###################################################
    #
    #   "find_string" Function is used to search a
    #   particular string in a .csv file. It returns
    #   string(s)'s Index and its count/occurrence.
    #
    #   "data" should be any .csv format file
    #
    #   It can search string in Row-wise as well as
    #   Column-wise. It will search Row-wise if selection
    #   is set to 0. While for 1 it will search Column-wise
    #   
    #   "number" is used to select particular row or 
    #   column in which want column-wise or row-wise
    #   search. 
    #   
    #   Example:
    # 
    #   Let's say we have a .csv file called "ABC" 
    #   which has 2 columns and 4 rows. Now we want 
    #   to search string "XYZ" in 2st column, row-wise
    #   search. So our setting should be:
    #   
    #   find_string(ABC,1,0,'XYZ') 
    #   number=1 as python starts from 0
    #
    #   It returns: 
    #               1.)Index of that string 
    #               2.)Its occurrence/count 
    #
    #   TIP: 
    #   To find string in whole document(.csv), just
    #   keep it in a loop, where "number" should be a 
    #   variable.
    ###################################################
    def find_string_sse(data,number,selection,input_string):

        data=np.array(data)
        count = 0
        index=[]

        if selection==0: # 0 is for row-wise search
            for i in range(len(data)):
                if str(input_string) in data[i,number]:
                    count += 1;
                    index.append(i)
            #print(count,index)
            return [count,index]

        elif selection==1: # 1 is for column-wise search
            for i in range(len(data[0,:])):
                if str(input_string) in data[number,i]:
                    count += 1;
                    index.append(i)
            #print(count,index)
            return [count,index]

        else:
            return print('Invalid Selection, Enter 0 for Row-wise search OR 1 for Column-wise search')
          
          
          
    ###################################################
    #   This function will read the generated .csv
    #   file and generate a baseline plot and return
    #   the read data. 
    ###################################################       
    def create_all_drives_baseline_plot(csv_name, file_dir
                                        , no_of_files, test_name_list):
        
        baseline_data= pandas.read_csv( r''+str(file_dir)+'\\'
                                   + str(csv_name)+'.csv'
                                   , header= None )
        
        [drive_counts, drive_indices]=SSE_Functions.find_string_sse(
                                        baseline_data
                                         , 0, 0, 'Drive#')
                                     
                                           
                                                
                                                
        baseline_data=np.array(baseline_data)
        
        # IOps
        iops_drive_no= [ baseline_data[drive_indices[i],0]
                        for i in range(no_of_files) ]

        iops_1st_test= [ baseline_data[drive_indices[i],1]
                        for i in range(no_of_files) ]

        iops_2nd_test= [ baseline_data[drive_indices[i],2]
                        for i in range(no_of_files) ]
                        
        [system_counts, system_indices]= SSE_Functions.find_string_sse(baseline_data, 0, 0, 'System')
        
        iops_system= baseline_data[system_indices[0],0]

        iops_1st_avg= baseline_data[system_indices[0],1]

        iops_2nd_avg= baseline_data[system_indices[0],2]
        
        
        summary_drive_list= [str(iops_system)]+iops_drive_no 
                             
        summary_1st_test_indi= [str(iops_1st_avg)] + iops_1st_test
        #summary_1st_test_indi

        summary_2nd_test_indi= [str(iops_2nd_avg)] + iops_2nd_test
        #summary_2nd_test_indi
        
        # Errors
        errors_1st_test= [ baseline_data[drive_indices[i],1]
                        for i in range(no_of_files,2*no_of_files) ]

        errors_2nd_test= [ baseline_data[drive_indices[i],2]
                        for i in range(no_of_files,2*no_of_files)]
                        
        
        errors_1st_avg= baseline_data[system_indices[1],1]

        errors_2nd_avg= baseline_data[system_indices[1],2]
        
                                
        summary_1st_test_indi_errors= [str(errors_1st_avg)] + errors_1st_test
        #summary_1st_test_indi

        summary_2nd_test_indi_errors= [str(errors_2nd_avg)] + errors_2nd_test
        
        # plot
        
        #x= np.linspace(0,no_of_files,no_of_files)
        #x_ticks= iops_drive_no    

        #%matplotlib inline
        #import matplotlib.pyplot as plt
        
        #fig1=plt.figure(1,figsize=(12,10))

        #plt.plot(x,iops_1st_test,'bo-', x,iops_2nd_test,'rs-' )

        #plt.xticks(x, x_ticks, rotation=90)
        #plt.grid(b=True, which='major', color='0.65',linestyle='-')
        #plt.grid(b=True, which='minor', color='0.65',linestyle='--')
        #plt.legend(test_name_list[1:],loc='best',bbox_to_anchor=(0.90, 1),
        #                                  bbox_transform=plt.gcf().transFigure)
        #plt.title('HBA cabled \nSA Basline Testing \nIOMeter Test Results by Drive ')
        #plt.ylim((0,float(max(iops_1st_test))+80))
        #plt.ylabel('IOps')
        #plt.xlim((0,no_of_files))
        
        #fig1.savefig(r''+str(file_dir)+ '\\'+ str(csv_name)+'_Plot_1.png')
        #plt.clf() 
        
        return [summary_drive_list, summary_1st_test_indi
                ,summary_2nd_test_indi, summary_1st_test_indi_errors
                , summary_2nd_test_indi_errors]
         

        
    ######################
    # Function will turn
    # test file's name
    # to all capital 
    # letters and extract
    # just the capacities
    # to qualify the file
    # names
    ######################
    def get_file_qualifiers(test_name_list):
        
        test_name_list_caps=[test_name_list[i].upper() 
                    for i in range(len(test_name_list)) ]
        
        file_qualifiers=[]
        temp_list_name= test_name_list_caps[1:]
        
        for j in range(len(temp_list_name)):
            temp = temp_list_name[j]
            for i in range(len(temp)):
                if temp[i] =='B':
                    last_index= i
                    break
            file_qualifiers.append(temp[:last_index+1])
        return [file_qualifiers, test_name_list_caps]  
        
        
       
    ######################
    # Function will take 
    # "file qualifiers" 
    # as inputs and returns
    # name of SSE files
    # with that qualifiers
    # It works for two 
    # files
    ######################
    def find_sse_with_qualifiers(test_name_list
                                 , file_list):
                                 
        [file_qualifiers, test_name_list_caps]= SSE_Functions.get_file_qualifiers(
                                             test_name_list)
        
        file_list =[file_list[i] for i in range(len(
                                          file_list)) 
                if 'SSE' in file_list[i] ]
        
        test_1_file= [file_list[i] for i in range(
                                    len(file_list))
                  if file_qualifiers[0] in file_list[i]]
        
        test_2_file= [file_list[j] for j in range(
                                    len(file_list)) 
                  if file_qualifiers[1] in file_list[j]]
        
        return [test_1_file, test_2_file, test_name_list_caps]

    
    
    ######################
    # This function will
    # extract indices of 
    # IOps and Errors  
    # data of all
    # drives for both the  
    # tests (i.e. Both 
    # files)
    ######################
    def get_indices_filedata(test_1_file, test_2_file
                             , test_name_list_caps, file_dir, test_name_list):
        import sys
        sys.path.insert(0,'C:\SSE Python scripts\Common Scripts')
        import report_functions
    
        rf= report_functions.Report_Functions
        
        test_1_data= pandas.read_csv( r''+str(file_dir)+'\\'
                                       + str(test_1_file[0])
                                       , skiprows=13
                                       , header= None )

        test_2_data= pandas.read_csv( r''+str(file_dir)+'\\'
                                           + str(test_2_file[0])
                                           , skiprows=19
                                           , header= None )
                                           
                                           
        ###############################
        # Finding indices of test name 
        # In this case, find indices of
        # 4K_67_33 and 512KB_0_100
        ###############################
        [test_1_counts, test_1_indices]= rf.find_string(
                                        test_1_data, 2, 0,
                                        test_name_list[1])
        
        # Finding IOps's Column number
        [test1_iops_counts, test1_iops_index]= rf.find_string(
                                               test_1_data, 0, 1,
                                               'IOps')

        # Finding Error's Column number
        [test1_e_counts, test1_error_index]= rf.find_string(
                                             test_1_data, 0, 1,
                                             'Errors')
                                             
        # Finding Number of disk allocated to each WORKER, Test 1
        [temp, test_worker_indices1]= rf.find_string(test_1_data,
                                                    0, 0,
                                                    'WORKER')
        test_1_no_of_disk = test_worker_indices1[1] - test_worker_indices1[0] - 1
                                            
                                                    
        # Finding indices of 2nd Test
        # In this case: 512KB_0_100_90s_10r
        [test_2_counts, test_2_indices]= rf.find_string(
                        test_2_data, 2, 0, test_name_list[2])
        
                                             
        # Finding Number of disk allocated to each WORKER, Test 1
        [temp, test_worker_indices2]= rf.find_string(test_2_data,
                                                    0, 0,
                                                    'WORKER')
                                                    
        test_2_no_of_disk = test_worker_indices2[1] - test_worker_indices2[0] - 1
        
        
        return [test_1_data,test_2_data
                , test_1_counts, test_1_indices
                , test1_iops_counts, test1_iops_index
                , test1_e_counts, test1_error_index
                , test_1_no_of_disk, test_2_no_of_disk
                , test_2_counts, test_2_indices]
                
                
                
    ######################            
    # From 1st Performance file            
    ######################
    # This function will
    # extract IOps and 
    # Errors data of all
    # drives for both the  
    # tests (i.e. Both 
    # files)
    ######################
    def get_performance_data(test_1_data,test_2_data
                , test_1_counts, test_1_indices
                , test1_iops_counts, test1_iops_index
                , test1_e_counts, test1_error_index
                , test_1_no_of_disk, test_2_no_of_disk
                , test_2_counts, test_2_indices, file_dir
                , test_1_file, test_2_file, serial_nos
                , sn_dic, vn_dic, mn_dic, keys_vn):
        

        ############################
        # For 1st Performance file
        ############################
        # To extract worker
        # numbers for 1st file
        ############################
        test_1_data= np.array(test_1_data)
        
        
        #####################
        # "test_1_no_of_disk"
        # will make fetching 
        # Drive data Dynamic 
        #####################
        test_1_phy_drive=[]
        for i in range(len(test_1_indices)):
            for j in range(1, test_1_no_of_disk+1):
                test_1_phy_drive.append(test_1_data[test_1_indices[i]+j,1])
        
        #print(test_1_phy_drive)

        
        ########################
        # Finding worker/physical drive nos to
        # cross check with 
        # Drive number from 
        # Individual drive's
        # names.
        ########################
        import re
        phy_drive_nos_1=[]
        for i in range(len(test_1_phy_drive)):
            
            #phy_drive= [int(s) for s in (test_1_phy_drive[i]).split() if s.isdigit()]
            phy_drive = re.findall(r'\d+', test_1_phy_drive[i])
            
            phy_drive_nos_1.append(int(phy_drive[0]))
            
        #print(phy_drive_nos_1)
        # 4k IOps data
        test_1_data= np.array(test_1_data)
 
        test_1_iops=[]
        test_1_errors=[]
        for i in range(len(keys_vn)):
        
            for j in range(int(len(phy_drive_nos_1)/test_1_no_of_disk)):
                
                for k in range(1, test_1_no_of_disk+1):
                    #print(j, k, (2*j)+k-1)
                    if int(phy_drive_nos_1[(2*j)+k-1]) == int(keys_vn[i]):

                        #1st test, IOps
                        temp1= test_1_data[test_1_indices[j]+k, test1_iops_index]
                        
                        temp1=float(np.asscalar(temp1))
                        test_1_iops.append(temp1)

                        #1st test, Errors
                        temp2= test_1_data[test_1_indices[j]+k,test1_error_index]
                        temp2=float(np.asscalar(temp2))
                        test_1_errors.append(temp2)

        #print(len(test_1_iops))
        from statistics import mean

        #1st test, IOps
        test_1_iops= [round(mean(test_1_iops),6)] + test_1_iops

        #1st test, IOps
        test_1_errors= [round(mean(test_1_errors),6)] + test_1_errors

        #print(final_1_iops, final_1_errors, len(final_1_iops))
        
            
        ######################            
        # From 2nd Performance file            
        ######################
        #####################
        # To extract worker
        # numbers for 2nd file
        #####################
        test_2_data= np.array(test_2_data)
        
        
        #####################
        # "test_2_no_of_disk"
        # will make fetching 
        # Drive data Dynamic 
        #####################
        test_2_phy_drive=[]
        for i in range(len(test_2_indices)):
            for j in range(1, test_2_no_of_disk+1):
                test_2_phy_drive.append(test_2_data[test_2_indices[i]+j,1])
            
        

        ########################
        # Finding worker nos to
        # cross check with                                
        # Drive number from 
        # Individual drive's
        # names.
        ########################
        phy_drive_nos_2=[]
       
        for i in range(len(test_2_phy_drive)):
            
            phy_drive = re.findall(r'\d+', test_2_phy_drive[i])
            
            phy_drive_nos_2.append(int(phy_drive[0]))
            

        # 512k IOps data
        test_2_data= np.array(test_2_data)
        #test_1_iops=[]

        test_2_iops=[]
        test_2_errors=[]
        for i in range(len(keys_vn)):
            for j in range(int(len(phy_drive_nos_2)/test_2_no_of_disk)):
                
                for k in range(1, test_2_no_of_disk+1):
                
                #print(wrkr_no)
                    if int(phy_drive_nos_2[(2*j)+k-1]) == int(keys_vn[i]):
                        
                        #2nd test, IOps
                        temp1= test_2_data[test_2_indices[j]+k,test1_iops_index]
                        temp1=float(np.asscalar(temp1))
                        test_2_iops.append(temp1)

                        #2nd test, Errors
                        temp2= test_2_data[test_2_indices[j]+k,test1_error_index]
                        temp2=float(np.asscalar(temp2))
                        test_2_errors.append(temp2)

        #print(len(test_1_iops))
        
        from statistics import mean

        #2nd test, IOps
        test_2_iops= [round(mean(test_2_iops),6)] + test_2_iops

        #2nd test, IOps
        test_2_errors= [round(mean(test_2_errors),6)] + test_2_errors

        #print(final_2_iops, final_2_errors, len(final_2_iops))
                
        
        # # Finding Time stamp info
        perf_1st= pandas.read_csv( r''+str(file_dir)+'\\'
                           + str(test_1_file[0])
                           , nrows=6
                           , header= None )

        time_stamp_1st_perf= perf_1st[0][5]
        #time_stamp_1st_perf

        perf_2nd= pandas.read_csv( r''+str(file_dir)+'\\'
                           + str(test_2_file[0])
                           , nrows=6
                           , header= None )
                           
        #time_stamp_2nd_perf
        time_stamp_2nd_perf= perf_2nd[0][5]
        
        perf_time_stamps= [time_stamp_1st_perf]+[time_stamp_2nd_perf]

        
        return [test_1_iops, test_1_errors
                , test_2_iops, test_2_errors
                , perf_time_stamps]
     

     
    ##########################
    # This function will
    # find degradation 
    # between individual_list
    # and performance_list
    # data.
    ##########################
    def find_degradation(individual_list, performance_list):
        degradation_list= [round(((float(individual_list[i]) 
         - float(performance_list[i]))/float(individual_list[i]))*100,1) for i in range(len(individual_list))]

        return degradation_list
    


    ##########################
    # This function will
    # write data in given
    # Worksheet row-wise or 
    # column-wise. "row"
    # "col" are row and column
    # from where we want to 
    # start writing out data.
    #
    # For row-wise:
    # rc_selection=0
    #
    # For column-wise:
    # rc_selection=1
    #
    ##########################    
    def write_excel_data(worksheet,data_list, row
                         , col, rc_selection, f_attributes):
       
        if rc_selection==0: #row wise
            
            for items in (data_list):
                worksheet.write(row, col, items, f_attributes)
                row+= 1
        elif rc_selection==1: #column wise
            
            for items in (data_list):
                worksheet.write(row, col, items, f_attributes)
                col+= 1
        else:
            
            print('Invalid input! \n'\
            '0: Row wise write\n 1: Column wise write')

            
            
    ##########################
    # Same function as above,
    # it will only convert values
    # to float with 2 decimal 
    # places. 
    ##########################   
    def write_excel_data_float(worksheet,data_list, row
                         , col, rc_selection, f_attributes):
       
        if rc_selection==0: #row wise
            
            for items in (data_list):
                worksheet.write(row, col, round(float(items),2), f_attributes)
                row+= 1
        elif rc_selection==1: #column wise
            
            for items in (data_list):
                worksheet.write(row, col, round(float(items),2), f_attributes)
                col+= 1
        else:
            
            print('Invalid input! \n'\
            '0: Row wise write\n 1: Column wise write')
    
    
    
    ##########################
    # To write Summary worksheet
    # in a Report
    ##########################    
    def write_summary(workbook, summary_ts, summary_td
                , summary_drive_list, summary_1st_test_indi
                , summary_2nd_test_indi, test_1_iops
                , test_2_iops, self_degradation_1st_test
                , self_degradation_2nd_test, degradation_1st_test
                , degradation_2nd_test, max_list, min_list
                , avg_list, std_list, summary_1st_test_indi_errors
                , summary_2nd_test_indi_errors, test_1_errors
                , test_2_errors):
                
        worksheet = workbook.add_worksheet('Summary') 
        
        # set width of Columns A thru J to 25
        worksheet.set_column('A:J', 25)
                
        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'yellow',
            'bold': True,
            'font_name':'Arial',
            'font_size':14})
            
        # Merge 3 cells over columns.
        worksheet.merge_range('A1:E1'
                              , 'VIBRATION TEST RESULTS'
                              , merge_format)
    
        # Add a bold format to use to highlight cells.
        bold_12 = workbook.add_format({'bold': True
                                      , 'font_name':'Arial'
                                      , 'font_size':12})

        regular_10 = workbook.add_format({'font_name':'Arial'
                                        , 'font_size':10
                                        , 'align': 'center'})
            
        # Font: Arial, Font_size:10 & Center Aligned
        regular_10.set_border(1)


        # Font: Arial, Font_size:10 & left Aligned
        regular_10_l = workbook.add_format({'font_name':'Arial'
                                    , 'font_size':10, 'align': 'left'})
        regular_10_l.set_border(1)
        
        row=0
        col=0
       
        #worksheet.write(row, col, '', bold_14)
        
        #title
        row+= 2
        worksheet.write(row, col, 'IOps Results', bold_12)

        
        # Time stamps
        row+= 1
        worksheet.write(row, col, summary_ts[0], regular_10_l)
        
        col=1
        SSE_Functions.write_excel_data(worksheet,summary_ts[1:], row
                             , col, 1, regular_10)

                             
        # Test description
        row+= 1
        col=0
        worksheet.write(row, col, summary_td[0], regular_10_l)
        
        
        # row from were test description start
        td_row = row
        #print(td_row)
        col=1
        SSE_Functions.write_excel_data(worksheet,summary_td[1:], row
                             , col, 1, regular_10)

        # Drives numbers
        row+= 1
        col=0
        SSE_Functions.write_excel_data(worksheet,summary_drive_list, row
                             , col, 0, regular_10_l)

                             
        # IOps data from Individual drives, 1st test
        col+= 1
        SSE_Functions.write_excel_data_float(worksheet,summary_1st_test_indi, row
                             , col, 0, regular_10)

                             
        # IOps data from Individual drives, 2nd test
        col+= 1
        SSE_Functions.write_excel_data_float(worksheet,summary_2nd_test_indi, row
                             , col, 0, regular_10)


        # IOps data from Performance file, 1st test
        col+= 1
        SSE_Functions.write_excel_data_float(worksheet,test_1_iops, row
                             , col, 0, regular_10)

                             
        # IOps data from Performance file, 2nd test
        col+= 1
        SSE_Functions.write_excel_data_float(worksheet,test_2_iops, row
                             , col, 0, regular_10)
        
        
        # Degradation data
        row+= len(test_2_iops)+1
        col= 0
        worksheet.write(row, col, 'IOps, % degradation,'\
                          ' baseline tests', bold_12)

                          
        # Drives numbers
        row+= 1
        col= 0
        SSE_Functions.write_excel_data(worksheet,summary_drive_list, row
                             , col, 0, regular_10_l)

                             
        # IOps data from Individual drives, 1st test
        col+= 1
        SSE_Functions.write_excel_data_float(worksheet,self_degradation_1st_test, row
                             , col, 0, regular_10)

                             
        # IOps data from Individual drives, 2nd test
        col+= 1
        SSE_Functions.write_excel_data_float(worksheet,self_degradation_2nd_test, row
                             , col, 0, regular_10)
        
        

        # IOps data from Performance file, 1st test
        col+= 1
        SSE_Functions.write_excel_data_float(worksheet,degradation_1st_test, row
                             , col, 0, regular_10)
    
        [workbook, worksheet] = SSE_Functions.add_conditional_formatting(0, workbook, worksheet, row, col, degradation_1st_test, -7, 7)                            
        
        # IOps data from Performance file, 2nd test
        col+= 1
        SSE_Functions.write_excel_data_float(worksheet,degradation_2nd_test, row
                             , col, 0, regular_10)

        
        [workbook, worksheet] = SSE_Functions.add_conditional_formatting(0, workbook, worksheet, row, col, degradation_2nd_test, -7, 7)                            
        
        
        #######################
        # Writing Stastical 
        # data
        #######################
        
        # High values
        row+= len(degradation_2nd_test)
        
        #Row where Statistics starts
        stats_row = row
        
        #Adding a top border
        temp_regular_font= regular_10
        temp_regular_font.set_border(1)
        col= 2
        SSE_Functions.write_excel_data(worksheet,max_list, row
                             , col, 1, temp_regular_font)
        
        [workbook, worksheet] = SSE_Functions.add_conditional_formatting(1, workbook, worksheet, row, col+1, max_list[1:], -7, 7)                            
        
        # Low values
        row+= 1
        SSE_Functions.write_excel_data(worksheet,min_list, row
                             , col, 1, regular_10)
        
        [workbook, worksheet] = SSE_Functions.add_conditional_formatting(1, workbook, worksheet, row, col+1, min_list[1:], -7, 7)                            
        
        # Average values
        row+= 1
        SSE_Functions.write_excel_data(worksheet, avg_list, row
                             , col, 1, regular_10)
        
        [workbook, worksheet] = SSE_Functions.add_conditional_formatting(1, workbook, worksheet, row, col+1, avg_list[1:], -7, 7)
        
        
        # Standard Deviation
        row+= 1
        SSE_Functions.write_excel_data(worksheet, std_list, row
                             , col, 1, regular_10)

        [workbook, worksheet] = SSE_Functions.add_conditional_formatting(1, workbook, worksheet, row, col+1, std_list[1:], -7, 7)
                
        
        # Errors
        # IO meter title
        row+= 2
        col= 0
        worksheet.write(row, col, 'IOMeter Errors', bold_12)
        
        # Drives numbers
        row+= 1
        SSE_Functions.write_excel_data(worksheet,summary_drive_list, row
                             , col, 0, regular_10_l)
        
        # Errors data from Individual drives, 1st test
        col+= 1
        SSE_Functions.write_excel_data_float(worksheet,summary_1st_test_indi_errors, row
                             , col, 0, regular_10)

        # Errors data from Individual drives, 2nd test
        col+= 1
        SSE_Functions.write_excel_data_float(worksheet,summary_2nd_test_indi_errors, row
                             , col, 0, regular_10)


        # Errors data from Performance file, 1st test
        col+= 1
        SSE_Functions.write_excel_data_float(worksheet,test_1_errors, row
                             , col, 0, regular_10)

        # Errors data from Performance file, 2nd test
        col+= 1
        SSE_Functions.write_excel_data_float(worksheet,test_2_errors, row
                             , col, 0, regular_10)
        
        return [td_row, stats_row]
        
        
        
    ###########################
    # Create Hi, Lo, Avg Chart
    ###########################
    # Not Used
    ###########################
    def create_degradation_plot(no_of_files, file_dir, csv_name
                                , iops_drive_no, test_name_list
                                , degradation_1st_test
                                , degradation_2nd_test):
                                
        x= np.linspace(0,no_of_files,no_of_files)
        x_ticks= iops_drive_no


        import matplotlib.pyplot as plt

        fig1=plt.figure(1,figsize=(12,10))

        plt.plot(x,degradation_1st_test[1:],'bo-', x,degradation_2nd_test[1:],'rs-' )

        plt.xticks(x, x_ticks, rotation=90)
        plt.grid(b=True, which='both', color='0.65',linestyle='-')
        #plt.grid(b=True, which='minor', color='0.65',linestyle='--')
        plt.legend(test_name_list[1:],loc='best',bbox_to_anchor=(0.9, 1),
                                           bbox_transform=plt.gcf().transFigure)

        plt.title('SSE on Rails and Blocks \nIOMeter Test Results by Drive ')
        #plt.ylim((0,float(max(degradation_1st_test[1:][1:]))+5))
        plt.ylim((-20,10))
        plt.ylabel('IOps % Degradation')
        #plt.axis('tight')
        plt.xlim((0,no_of_files))
        fig1.savefig(r''+str(file_dir)+ '\\'+ str(csv_name)+'_Plot_2.png')
        plt.clf() 
        
        
    #########################
    # Creating 
    # Hi, Lo, Avg Chart in 
    # a given Workbook.
    #########################
    def create_hi_lo_avg_chart(td_row, stats_row, workbook):
        
        worksheet = workbook.add_worksheet('Hi, Lo, Avg Chart')

        chart = workbook.add_chart({'type': 'stock'})
        
        worksheet.set_column('B:L', 25) #column width
        
        # For degradation limit
        line_data_x= [-1,1]
        line_data_y= [7,7]
        
        # writing above data in a worksheet
        worksheet.write_column('Y2', line_data_x)
        worksheet.write_column('Z2', line_data_y)
        
        #adding a chart
        line_chart = workbook.add_chart({'type': 'line'})

        # Configure a series with a secondary axis
        line_chart.add_series({
                'name': '7% maximum allowed degradation',
                #'categories': '= Hi, Lo, Avg Chart!$Y$2:$Y$3',
                'values': '=Hi, Lo, Avg Chart!$Z$2:$Z$3',
                #'backward': 0.5,
                'line': {'color':'red','width': 1.5},
                              })

        #chart.set_y2_axis({'name': 'Degradation value'})
        chart.combine(line_chart)
        chart.set_size({'x_scale': 1.5, 'y_scale': 1.5})

        # Add a series for each of the High-Low-Close columns.
        chart.add_series({
            'line': 'o',    
            'name': 'High',
            #'trendline': {'type': 'moving_average'},  
            #'fill':   {'none': False},
            'categories': '=Summary!$B$' +str(td_row+1)
                            + ':$C$'+str(td_row+1),
                
            'values': '=Summary!$D$' +str(stats_row+1)
                        +':$E$'+str(stats_row+1),
            'line':   {'none': True},    
            #'line':       {'color': 'red'},
            'marker': {
                        'type': 'diamond',
                        'border': {'color': 'red'},
                        'fill':   {'color': 'red'},
                    }
                
            
                        })

                        
        chart.add_series({
            'name': 'Low',
            'categories': '=Summary!$B$' +str(td_row+1)
                            + ':$C$'+str(td_row+1),
                
            'values':     'Summary!$D$' +str(stats_row+2)
                        +':$E$'+str(stats_row+2),
                
            'line':   {'none': True},
            #'line':       {'color': 'blue'},
            'marker': {
                        'type': 'square',
                        'border': {'color': 'blue'},
                        'fill': {'color':'blue'}
                      }
                          })

                          
        chart.add_series({
            'name': 'Average',
            'categories': '=Summary!$B$' +str(td_row+1)
                            + ':$C$'+str(td_row+1),
                
            'values':     'Summary!$D$' +str(stats_row+3)
                        +':$E$'+str(stats_row+3),
                
            'line':   {'none': True},
            #'line':       {'color': 'green'},
            'marker': {
                        'type': 'square',
                        'border': {'color': 'green'},
                        'fill': {'color':'green'}
                        #'set_y_axis':
                      }
                        })
                        
                        
        chart.set_x_axis({'label_position': 'low'
                         , 'position_axis': 'on-ticks'
                         , 'min': 1, 'max': 2
                         })
                            
                            
        chart.set_y_axis({'name': 'IOps % Degradation',
                         'min': -10, 'max': 10})

                         
        chart.set_title({'name':'SSE on Rails and Blocks '\
                         '\nIOMeter Test Results, High/Low/Average'})

        worksheet.insert_chart('B2', chart)
        
        
        
    #########################
    # Creating 
    # All Drives Vibe Chart
    #########################
    def create_all_drives_vibe(td_row, stats_row, workbook
                                , no_of_files, summary_td
                                , keys_vn, values_vn
                                , values_mn, values_sn ):
                                
        worksheet = workbook.add_worksheet('All Drives Vibe')
        regular_10 = workbook.add_format({'font_name':'Arial'
                                    , 'font_size':10, 'align': 'center'})
        
        # Font: Arial, Font_size:10 & Center Aligned
        worksheet.set_column('B:L', 25)
                
        # Font: Arial, Font_size:10 & Center Aligned
        regular_10.set_border(1)
        col = 1
        row = 24
        
        # writing drive number, vendor name, model number 
        # and Serial number
        SSE_Functions.write_excel_data(worksheet, keys_vn, row, col, 0, regular_10)
        SSE_Functions.write_excel_data(worksheet, values_vn, row, col+1, 0, regular_10)
        SSE_Functions.write_excel_data(worksheet, values_mn, row,col+2, 0, regular_10)
        SSE_Functions.write_excel_data(worksheet, values_sn, row, col+3, 0, regular_10)

        
        chart = workbook.add_chart({'type': 'line'})
        chart.set_size({'x_scale': 1.5, 'y_scale': 1.5})

        chart.add_series({
            #'line': 's',    
            'name': str(summary_td[1]),
            #'trendline': {'type': 'moving_average'},  
            #'fill':   {'none': False},
            'categories': '=Summary!$A$' +str(stats_row-no_of_files+1)
                            + ':$A$'+str(stats_row),
                
            'values': '=Summary!$D$' +str(stats_row-no_of_files+1)
                        +':$D$'+str(stats_row),
                                         
            #'line':   {'none': True},    
            'line':       {'color': 'pink'},
            'marker': {
                        'type': 'square',
                        'border': {'color': 'pink'},
                        'fill':   {'color': 'pink'},
                    }
                        })

        chart.add_series({
            #'line': 's',    
            'name': str(summary_td[2]),
            #'trendline': {'type': 'moving_average'},  
            #'fill':   {'none': False},
            'categories': '=Summary!$A$' +str(stats_row-no_of_files+1)
                            + ':$A$'+str(stats_row),
                
            'values': '=Summary!$E$' +str(stats_row-no_of_files+1)
                        +':$E$'+str(stats_row),
                                         
            #'line':   {'none': True},    
            'line':       {'color': 'blue'},
            'marker': {
                        'type': 'square',
                        'border': {'color': 'blue'},
                        'fill':   {'color': 'blue'},
                    }
                        })

            
        chart.set_title({'name':'SSE on Rails and Blocks '\
                         '\nIOMeter Test Results by Drive'})

        chart.set_y_axis({'name': 'IOps % Degradation',
                         'min': -20, 'max': 10})

        chart.set_x_axis({'label_position': 'low'
                          , 'position_axis': 'on_tick'
                          , 'major_gridlines': {
                                    'visible': True,
                                    'line': {'width': 0.5}
                        }})
                        
                        
        worksheet.insert_chart('B2', chart)
    
    
    
    #########################
    # Creating 
    # Baseline data
    #########################
    def create_baseline(workbook, summary_ts, summary_td
                        , summary_drive_list, summary_1st_test_indi
                        , summary_2nd_test_indi
                        , summary_1st_test_indi_errors
                        , summary_2nd_test_indi_errors ):
    
        
        worksheet = workbook.add_worksheet('Baseline')
        worksheet.set_column('A:J', 25)

        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'yellow',
            'bold': True,
            'font_name':'Arial',
            'font_size':14})

        # Merge 3 cells over columns.
        worksheet.merge_range('A1:C1', 'BASELINE TEST RESULTS', merge_format)
        worksheet.merge_range('A2:C2', 'HBA cabled', merge_format)

        # Add a bold format to use to highlight cells.
        bold_14 = workbook.add_format({'bold': True, 'font_name':'Arial'
                                    , 'font_size':14})


        bold_12 = workbook.add_format({'bold': True, 'font_name':'Arial'
                                    , 'font_size':12})

        regular_10 = workbook.add_format({'font_name':'Arial'
                                    , 'font_size':10, 'align': 'center'})

        # Font: Arial, Font_size:10 & Center Aligned
        regular_10.set_border(1)


        # Font: Arial, Font_size:10 & left Aligned
        regular_10_l = workbook.add_format({'font_name':'Arial'
                                    , 'font_size':10, 'align': 'left'})
        regular_10_l.set_border(1)


        row=0
        col=0
        #worksheet.write(row, col, 'BASELINE TEST RESULTS', bold_14)

        row+=1
        #worksheet.write(row, col, 'HBA cabled', bold_14)


        #title
        row+=2

        #bold_12.set_bottom(1)
        worksheet.write(row, col, 'IOps Results', bold_12)

        #format = workbook.add_format()


        row+= 1
        worksheet.write(row, col, summary_ts[0], regular_10_l)
        
        # Time stamps
        col=1
        SSE_Functions.write_excel_data(worksheet,summary_ts[1:3], row
                             , col, 1, regular_10)

        # Test description
        row+= 1

        td_row_indi = row
        col=0
        worksheet.write(row, col, summary_td[0], regular_10_l)

        col=1
        SSE_Functions.write_excel_data(worksheet,summary_td[1:3], row
                                      , col, 1, regular_10)

        # Drives numbers
        row+= 1
        col=0
        SSE_Functions.write_excel_data( worksheet
                                      , summary_drive_list, row
                                      , col, 0, regular_10_l)

        # IOps data from Individual drives, 1st test
        col+= 1
        SSE_Functions.write_excel_data_float( worksheet
                                            , summary_1st_test_indi
                                            , row
                                            , col
                                            , 0
                                            , regular_10)

        # IOps data from Individual drives, 2nd test
        col+= 1
        SSE_Functions.write_excel_data_float( worksheet
                                            , summary_2nd_test_indi
                                            , row
                                            , col
                                            , 0
                                            , regular_10)


        #Errors
        # IO meter title
        row= row + 1 + len(summary_2nd_test_indi)
        col= 0

        worksheet.write(row, col, 'IOMeter Errors', bold_12)



        # Drives numbers
        row+= 1
        SSE_Functions.write_excel_data( worksheet
                                      , summary_drive_list
                                      , row
                                      , col
                                      , 0
                                      , regular_10_l)

        # Errors data from Individual drives, 1st test
        col+= 1

        SSE_Functions.write_excel_data_float( worksheet
                                            ,summary_1st_test_indi_errors
                                            , row
                                            , col
                                            , 0
                                            , regular_10)
    
    
        # Errors data from Individual drives, 2nd test
        col+= 1
        SSE_Functions.write_excel_data_float( worksheet
                                            , summary_2nd_test_indi_errors
                                            , row
                                            , col   
                                            , 0
                                            , regular_10)

        return td_row_indi
        
        
        
    #########################
    # Creating 
    # All Drives Baseline
    #########################
    def create_all_drives_baseline(workbook, summary_td, td_row_indi
                                   , no_of_files, keys_vn, values_vn
                                   , values_mn, values_sn ):

        worksheet = workbook.add_worksheet('All Drives Baseline')
        
        regular_10 = workbook.add_format({'font_name':'Arial'
                                        , 'font_size':10
                                        , 'align': 'center'})
        
        worksheet.set_column('B:L', 25)
        
        # Font: Arial, Font_size:10 & Center Aligned
        regular_10.set_border(1)
        
        col = 1
        row = 24
        
        # writing drive number, vendor name, model number 
        # and Serial number
        SSE_Functions.write_excel_data(worksheet, keys_vn, row, col, 0, regular_10)
        SSE_Functions.write_excel_data(worksheet, values_vn, row, col+1, 0, regular_10)
        SSE_Functions.write_excel_data(worksheet, values_mn, row,col+2, 0, regular_10)
        SSE_Functions.write_excel_data(worksheet, values_sn, row, col+3, 0, regular_10)

        

        chart = workbook.add_chart({'type': 'line'})
        chart.set_size({'x_scale': 1.5, 'y_scale': 1.5})

        chart.add_series({
            #'line': {'width': 0.1},    
            'name': str(summary_td[1]),
            #'trendline': {'type': 'moving_average'},  
            #'fill':   {'none': False},
            'categories': '=Baseline!$A$' +str(td_row_indi+3)
                            + ':$A$'+str(td_row_indi+no_of_files+2),
                
            'values': '=Baseline!$B$' +str(td_row_indi+3)
                        +':$B$'+str(td_row_indi+no_of_files+2),
                                         
            #'line':   {'none': True},    
            'line':       {'color': 'pink'},
                
            #'data_labels': {'series_name': True, 'position': 'above'},    
            'marker': {
                        'type': 'square',
                        'border': {'color': 'pink'},
                        'fill':   {'color': 'pink'},
                      }
                      
                        })

        chart.add_series({
        
            #'line': {'width': 0.1},    
            'name': str(summary_td[2]),
            #'trendline': {'type': 'moving_average'},  
            #'fill':   {'none': False},
            'categories': '=Baseline!$A$' +str(td_row_indi+3)
                            + ':$A$'+str(td_row_indi+no_of_files+2),
                
            'values': '=Baseline!$C$' +str(td_row_indi+3)
                        +':$C$'+str(td_row_indi+no_of_files+2),
                                          
            #'line':   {'none': True},    
            'line':       {'color': 'blue'},
            #'data_labels': {'series_name': True, 'position': 'above'},   
            'marker': {
                        'type': 'square',
                        'border': {'color': 'blue'},
                        'fill':   {'color': 'blue'},
                      }
                      
                        })

        chart.set_title({'name':'HBA cabled \nSA Basline Testing'\
                         '\nIOMeter Test Results by Drive '})
                        
        chart.set_y_axis({'name': 'IOps',
                         'min': 0})
                         

        chart.set_x_axis({'label_position': 'low'
                          , 'position_axis': 'on_tick'
                          , 'min': 1
                          , 'major_gridlines': {
                                    'visible': True,
                                    'line': {'width': 0.5}
                        }})

        worksheet.insert_chart('B2', chart)
        
        
        
    ############################################
    # CONDITIONAL FORMATING
    #
    # This function will change 
    # the font color conditionally.
    #
    #   RED:
    #    if "red_value"
    #    is greater than and equal to value(>=) 
    #    Cell value.
    #
    #   BLUE:
    #    if "blue_value"
    #    is less than and equal to value(<=) 
    #    Cell value.
    #
    #   "selection" used to select row-wise
    #    or column-wise formatting
    #   
    #    if 
    #       selection = 0: Row Wise formatting
    #       selection = 1: Column Wise formatting
    #                      with same Font color 
    #                      scheme
    ############################################
    def add_conditional_formatting(selection, workbook, worksheet, row, col, list, blue_value, red_value):
        
        ssef = SSE_Functions
        
        list_length = len(list)
        
        # For RED FONT
        red_format = workbook.add_format({
                                            'font_color': 'red'
        
                                        })
        
        # For BLUE FONT
        blue_format = workbook.add_format({
                                            'font_color': 'blue'
                                         })
        if selection == 0: # ROW_WISE                                 
            worksheet.conditional_format('$'+str(ssef.rank(col+1))+ '$' + str(row+1) + ':$' +str(ssef.rank(col+1))+ '$' +str(row+list_length),

                                            {
                                                'type': 'cell',
                                                'criteria': '>=',
                                                'value': red_value,
                                                'format': red_format,
                                            }
                                        )
            
            worksheet.conditional_format('$'+str(ssef.rank(col+1))+ '$' + str(row+1) + ':$' +str(ssef.rank(col+1))+ '$' +str(row+list_length),
                                            {
                                                'type': 'cell',
                                                'criteria': '<=',
                                                'value': blue_value,
                                                'format': blue_format,
                                            }
                                        )
        elif selection == 1: # COLUMN_WISE                                 
            worksheet.conditional_format('$'+str(ssef.rank(col+1))+ '$' + str(row+1) + ':$' +str(ssef.rank(col+list_length))+ '$' +str(row+1),

                                            {
                                                'type': 'cell',
                                                'criteria': '>=',
                                                'value': red_value,
                                                'format': red_format,
                                            }
                                        )
            
            worksheet.conditional_format('$'+str(ssef.rank(col+1))+ '$' + str(row+1) + ':$' +str(ssef.rank(col+list_length))+ '$' +str(row+1),
                                            {
                                                'type': 'cell',
                                                'criteria': '<=',
                                                'value': blue_value,
                                                'format': blue_format,
                                            }
                                        )
        return [workbook, worksheet]                        
    
    ##################################
    # This Function will return an
    # Alphabet related to its Numbers
    #
    # Example rank(1) = A
    #         rank(26) = Z
    #
    # Note that number should be less
    # or equal to 26. Or else it will
    # throw a KeyError
    ##################################
    def rank(x):
        import string
        d = dict((n%26+1,letr) for n,letr in enumerate(string.ascii_letters[0:52]))
        return d[x]     
    