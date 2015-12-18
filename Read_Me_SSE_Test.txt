####################################################
#                                                  #
#   Instructions to run SSE scripts in Python      #
#                                                  #
#   Author: Zankar Sanghavi                        #
#                                                  #
#   © Dot Hill Systems Corporation                 #
#                                                  #
####################################################



####################################################
#                                                  #
#            Installation and running              #
#						   #
#  If you already installed Python 3 and Packages  #
#       then you can start from step:3  	   #
#                                                  #
####################################################

1.) Before running, you need to Install Anaconda Python 3
    from Continuum website or you can also use a Setup file 
    stored in "Setup Files" folder.

    Path: meng:\Hard Drive\Zankar Sanghavi\Setup Files\Anaconda3-2.2.0-Windows-x86_64

2.) Once you install Python, you need to install 3 Librabies
    or Packages (comtypes-1.1.1, openpyxl-2.2.5 and 
    XlsxWriter-0.7.3) to support these scripts. 
    
    To install a package:
    
    a.) Copy and Paste "Packages" in your local directory.
    
    b.) Open Command Prompt from Start Menu
    
    c.) Go to that directory by using this command:
    
        cd Absolute Path of Package
        
        For example:
        
            cd C:\Packages\comtypes-1.1.1
                
    d.) Then enter:
                
            ipython setup.py install
            
    Repeat process from Step (c) for each Package.
    
3.) Now you are ready to run a Report Automation script for SSE 
    Test. Change the path to Specific path and enter. 
    
    ipython.

4.) Once you are in iPython, you can simply run script by entering 
    this command:
    
    run main_sse.py

5.) Then you will need to Enter full path of SSE files (i.e. Performance 
    file) and name of Report(s).


NOTE: In order to generate a Report, directory structure should be:

	i.) Baseline files should be in a sub-directory with User defined name.
	
	iI.) Sub-Directory/Directories should contain all baselines with "Serial Number_sa_baseline" format.

	    	Example: XXXX01234_sa_baseline.csv

	iii.) Single "sn_multiple_drives.txt" file.

	iv.) There can be only two SSE/Performance files in that folder.	

	v.) Both SSE/Performance files must have "SSE" and test name string included in File name. 

	     	Example: SSE_4KB_67_33_XXXX0123.csv (i.e. it should SSE & 4KB (All capitals).

	vi.) Test names from Baseline should exactly match, Test names in Performance files. They are
	     case sensitive. 




####################################################
#                                                  #
#            Tips and Tricks for the User          #
#                                                  #
####################################################

1.) To change the directory, you can just Copy path from Address bar 
    and use right mouse click to paste it in a Command prompt window.

	Example: cd path; cd C:\SSE_Python\SSE

2.) After you get in to iPython, you can Browse files by entering 
    initial characters and then pressing Tab.
    
    Example:
    
    run  main_s -> Tab -> run main_sse.py 

	
3.) For files path you have to make sure that file exits, or a simple way
    get a path is to Drag and drop the file on Command Prompt window. You 
    can apply this where ever it says Enter path of a File.



####################################################
#                                                  #
#                   Know Issues                    #
#                                                  #
####################################################    

None, till 10/22/2015
