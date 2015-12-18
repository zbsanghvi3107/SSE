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


class Write_Single_Report:
    ########################################
    #  This function will Generate a
    #  a extract data from Baseline files
    #  as well as from Performance
    #  or SSE files and Generate a single 
    #  report.
    ########################################
    def Generate_Single_Report(file_dir, file_list, file_dir_perf, file_list_perf, csv_name):
        
        import pandas
        import os
        import numpy as np
        import xlsxwriter
        import sse_functions
        
        
        # accessing "sse_functions" fuction
        # from "SSE_Functions.py" file
        ssef= sse_functions.SSE_Functions
        
        ###################################
        # This function will extract 
        # Vendor name, Serial Number, and
        # Model Number from 
        # "sn_multiple_drives.txt" and 
        # relate to Physical Drive Number.
        ###################################
        [vn_dic, sn_dic, mn_dic] = ssef.generate_sn_mn_dictionary(file_dir_perf
                                                                  , file_list_perf)
                                                                  
                                                                  

        ###################################
        #  Extracting required data from
        #  all "baseline" .csv files.
        #  
        #  Moreover, it will use Serial
        #  numbers from generated from above 
        #  function to Sort Files as per 
        #  their drive numbers in Ascending
        #  order. 
        #
        #  It will also assign drive number 
        #  related to Serial Number.
        ###################################
        [test_name_list, final_time_stamps
         , iops_final_list, errors_final_list
         , no_of_files,serial_nos] = ssef.get_data_baseline(file_list
                                                           , file_dir
                                                           , sn_dic
                                                           , mn_dic)

                                                           
        ##########################################
        # This function will extract a sorted
        # list of Serial number, Vendor Name,
        # Model Number aligned with Drive Numbers
        ###########################################
        [values_sn
          , values_vn
          , values_mn
          , keys_vn] = ssef.extract_present_data(sn_dic
                                                , vn_dic
                                                , mn_dic
                                                , serial_nos)                                            
        
        
        ###################################
        #  Writing a new .csv from
        #  all "baseline" .csv files and
        #  storing in same directory
        # temporarily.
        ###################################    
        ssef.create_csv_sse(csv_name, file_dir, test_name_list
                            , final_time_stamps, iops_final_list
                            , errors_final_list)

                              
                              
        ###################################
        # This function will read new 
        # baseline.csv file and provide
        # data to write in Excel file.
        ###################################                        
        [summary_drive_list, summary_1st_test_indi
        , summary_2nd_test_indi, summary_1st_test_indi_errors
        , summary_2nd_test_indi_errors] = ssef.create_all_drives_baseline_plot(csv_name
                                                      , file_dir
                                                      , no_of_files
                                                      , test_name_list)
                                            
               

        # Using qualifiers get names of Performance files
        [test_1_file
        , test_2_file
        , test_name_list_caps] = ssef.find_sse_with_qualifiers(
                                                      test_name_list
                                                    , file_list_perf)

                                                    
                                                    
        # Get data of Performance files and get required indices                                            
        [test_1_data,test_2_data
         , test_1_counts, test_1_indices
         , test1_iops_counts, test1_iops_index
         , test1_e_counts, test1_error_index
         , test_1_no_of_disk, test_2_no_of_disk
         , test_2_counts, test_2_indices] = ssef.get_indices_filedata(
                                                test_1_file, test_2_file
                                                , test_name_list_caps
                                                , file_dir_perf
                                                , test_name_list)

                  
                  
        # Using Indices extract IOps and Error data from both files                                                
        [test_1_iops, test_1_errors
         , test_2_iops, test_2_errors
         , perf_time_stamps] = ssef.get_performance_data(
                                              test_1_data,test_2_data
                                            , test_1_counts
                                            , test_1_indices
                                            , test1_iops_counts
                                            , test1_iops_index
                                            , test1_e_counts
                                            , test1_error_index
                                            , test_1_no_of_disk
                                            , test_2_no_of_disk
                                            , test_2_counts
                                            , test_2_indices
                                            , file_dir_perf
                                            , test_1_file
                                            , test_2_file
                                            , serial_nos
                                            , values_sn
                                            , values_vn
                                            , values_mn
                                            , keys_vn)                                                

                                            
                                            
        # Time stamp for Summary worksheet
        summary_ts= final_time_stamps + perf_time_stamps

        # Test description for Summary worksheet
        summary_td= test_name_list + test_name_list[1:]  
        #summary_td

        # degradation between individual and combined
        # data from individual csv vs performance file

        # Self degradation 1st test
        self_degradation_1st_test= ssef.find_degradation(
            summary_1st_test_indi, summary_1st_test_indi)
        #print(self_degradation_1st_test)

        # Self degradation 2nd test
        self_degradation_2nd_test= ssef.find_degradation(
            summary_2nd_test_indi, summary_2nd_test_indi)
        
        ################
        # For 1st test
        ################
        #print('Summary: ' + str(len(summary_1st_test_indi)))
        #print('Test IOps: ' + str(len(test_1_iops)))
        
        degradation_1st_test= ssef.find_degradation(
            summary_1st_test_indi, test_1_iops)
        #print(degradation_1st_test)

        ################        
        # For 2nd test
        ################
        degradation_2nd_test= ssef.find_degradation(
            summary_2nd_test_indi, test_2_iops)


        #################################
        # For statistical calculations   
        #################################    
        
        # High(maximum) degradation values
        max_list= [round(float(np.max(degradation_1st_test[1:])),1)]+[round(float(np.max(degradation_2nd_test[1:])),1)] 
        max_list= ['High'] + max_list

        #Low(minimum) degradation values
        min_list= [round(float(np.min(degradation_1st_test[1:])),1)]+[round(float(np.min(degradation_2nd_test[1:])),1)] 
        min_list= ['Low'] + min_list

        # Average of degradation values
        avg_list= [round(float(np.mean(degradation_1st_test[1:])),1)]+[round(float(np.mean(degradation_2nd_test[1:])),1)] 
        avg_list= ['Average'] + avg_list

        # Standard deviation of degradation values
        std_list= [round(float(np.std(degradation_1st_test[1:])),1)]+[round(float(np.std(degradation_2nd_test[1:])),1)] 
        std_list= ['1 σ'] + std_list


            
        # Writing an empty Excel file, name given by User.
        workbook = xlsxwriter.Workbook(r''+str(file_dir)+
                                       '\\' +str(csv_name)
                                       + '.xlsx')

           
           
           
        #################################
        # Writing a "Summary" worksheet
        # from the extracted data. This 
        # data is extracted from 
        # Performance files.
        #################################
        [td_row, stats_row] = ssef.write_summary(workbook
                        , summary_ts, summary_td
                        , summary_drive_list, summary_1st_test_indi
                        , summary_2nd_test_indi, test_1_iops
                        , test_2_iops, self_degradation_1st_test
                        , self_degradation_2nd_test, degradation_1st_test
                        , degradation_2nd_test, max_list, min_list
                        , avg_list, std_list, summary_1st_test_indi_errors
                        , summary_2nd_test_indi_errors, test_1_errors
                        , test_2_errors)

                        
                        
        #################################
        # Using data from "Summary"
        # worksheet, generating a 
        # Stock chart in a new worksheet
        # called "Hi, Lo Avg Chart".
        #################################                
        ssef.create_hi_lo_avg_chart(td_row, stats_row, workbook)


        #################################
        # Using data from "Summary"
        # worksheet, generating a 
        # Line chart in a new worksheet
        # called "All Drives Vibe".
        #################################
               
        # Appending Title to the front
        
        # For Drive Numbers
        keys_vn1 = ['Drive No.'] + keys_vn
        
        # For Vendor Name
        values_vn1 = ['Vendor'] + values_vn
        
        # For Model Number
        values_mn1 = ['Model Number'] + values_mn
       
        # For Serial Number
        values_sn1 = ['Serial Number'] + values_sn
        
        
        # Write and create All Drives Vibe Worksheet.
        ssef.create_all_drives_vibe(td_row, stats_row, workbook
                                    , no_of_files, summary_td
                                    , keys_vn1, values_vn1
                                    , values_mn1, values_sn1 )

                       
                       
        #################################
        # Writing a "Baseline" worksheet
        # from the extracted data. This 
        # data is extracted from
        # Individual drives data.
        #################################                  
        td_row_indi= ssef.create_baseline(workbook
                                    , summary_ts, summary_td
                                    , summary_drive_list
                                    , summary_1st_test_indi
                                    , summary_2nd_test_indi
                                    , summary_1st_test_indi_errors
                                    , summary_2nd_test_indi_errors )


                                    
        #################################
        # Using data from "Baseline"
        # worksheet, generating a 
        # Line chart in a new worksheet
        # called "All Drives Baseline".
        #################################                            
        ssef.create_all_drives_baseline(workbook
                                        , summary_td, td_row_indi
                                        , no_of_files, keys_vn1
                                        , values_vn1
                                        , values_mn1
                                        , values_sn1 )
                

                
        workbook.close() 

        # removing temporarily generated .csv file
        os.remove(r'' +str(file_dir)+'\\' + str(csv_name)+'.csv')                                            