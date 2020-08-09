# Data-Analysis-AIT


## Author :  Mohamed El Amine OUIS


## Description
Main Code to analyse data from **Oscilator CSV files** containings **Tension** and **Current** Data for specific time period.
Each channel has it's own CSV file :
- channel 1 : Tension (V) [..._Ch1.csv]
- channel 2 : Current (A) [..._Ch2.csv]

The output files are :
- **Excel Sheet** containing all data and charts + info.
- **Figures images** for each channel with Delta time (ms).

### Release Date    : 08-08-2020

### Actual Version  : 1.1

### Changelog:

#### v1.1 :
- Added **config.ini** file to set all default parameters inside. The configurable parameters are :
    - N : filter window size
    - processed_dir : Directory to store the generated files : Excel + Figures
    - calibration_value_ch1 : Tension physical value (V) = Tension value * Calibration_value_ch1
    - calibration_value_ch2 : Current physical value (A) = Current value * Calibration_value_ch2
    - threshold_low_ch1 : min threshold for Tension : used to calculate delta time
    - threshold_high_ch1 : max threshold for Tension : used to calculate delta time
    - threshold_low_ch2 : min threshold for Current : used to calculate delta time
    - threshold_high_ch2 : max threshold for Current : used to calculate delta time
- Enhenced Console Display.
- Fixed file name in figure title.
- Added plt.close() after saving the figure to clear the memory.
- Other changes.


#### v1.0 :
- Select Working directory where the CSV files are stored
- Creation of new directory to save the processed data : Excel sheets + images
- The script lists all 'CSV' files in the working directory, each one with its own index
- The user select which indexes to process. The selection is possible using these methods:
    - One File : input the file index
    - Multiple Files (option 1) : input files indexes and separate wetween them with '.' or ',' or ';'
    - Multiple Files (option 2) : input idx_start and idx_end separated by ':' to get all files having indexes between them
    - All File : just input 'all' or 'a'
- It is possible to process each file seprately, or as a combined channels (Ch1_Tension + Ch2_Current)
- The implemented filter used to filter the data is SMA : Simple Moving Average with a windows Size = N (fixed in the main)
    
- **The process of combined files ch1 + ch2 :**
    - The files list will be displayed using same index for a combined files (ch1 + ch2)
    - Select files to process
    - For each File :
        - Create new folder to save Excel + images
        - Start by creating two (2) CSV_Object instances to load the data from CSV file
        - For each instance ch1 and ch2 apply these methods/functions :
            - apply_moving_average()
            - calculate_delta_time() for both raw and filtred data
            - create_plot_data() for both raw and filtred data
        - Generate combined exel file with graphs using this function :
            - generate_excel_combined()

- **The process of all files separetly :**
    - Select files indexes to process
    - For each File :
        - Create new folder to save Excel + images
        - Start by creating a CSV_Object instance to load the data from CSV file
        - apply_moving_average()
        - calculate_delta_time() for both raw and filtred data
        - create_plot_data() for both raw and filtred data
        - Generate exel file with graphs using this method : generate_excel()
     

# Installation

### Requirement
**Python3** : To run the code

Check installation : ```$ python --version```

Download link : [https://www.python.org/downloads]

**Git + pip/pip3 + virtualenv** : To install the code

Check installation :
```sh
$ git --version
$ pip --version             or          $ pip3 --version
$ virtualenv --version
```

Git Download Link : [https://git-scm.com/downloads]

Install virtualenv : ```$ pip install virtualenv```


### 1. Clone repository :
```sh
$ git clone https://github.com/OUIS-dev/Data-Analysis-AIT.git
```

### 2. Change directory to the new downloaded directory :
```sh
$ cd Data-Analysis-AIT
```

### 3. Create new virtual envirement : 
```sh
$ virtualenv env
```

### 4. Activate the virtualenv :
- Linux or Mac :
```sh
$ source env/bin/activate
```
- Windows CMD :
```sh
$ env\Scripts\activate.bat
```

### 5. install required python packages :
```sh
(env)       $ pip install -r requirement.txt
```

### 6. Run the script :
```sh
(env)       $ cd src
(env)       $ python main.py
```

# Usage Example
```sh
(env) [ouis@ouis-pc src]$ python main.py 

Script Started ...

Enter 'CSV' Directory : /home/ouis/Documents/Dev/data_analysis_ait/data_samples/

--------------------------------------------   Combined Option   ---------------------------------------------

Process the two channels combined [yes or 1]
Process Each channel file one by one [no, 0]
Please select an option : 1


-----------------------------------------------   Files List   -----------------------------------------------

Directory : /home/ouis/Documents/Dev/data_analysis_ait/data_samples/

index : [0] --- File Name : 20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm
index : [1] --- File Name : 20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm


--------------------------------------------   Index Selection   ---------------------------------------------

Info :
One File : input the file index
Multiple Files (option 1) : input files indexes and separate between them with '.' or ',' or ';'
Multiple Files (option 2) : input idx_start and idx_end separated by ':' to get all files having indexes between them
All File : just input 'all' or 'a'

input index(es) : 0


---------------------------------------------   Selected file   ----------------------------------------------

index : [0] --- File Name : ('20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm_Ch1.csv', '20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm_Ch2.csv')



[1/1] Processing File : 20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm                                                   


----------------------------------------------------------------------------------------------------

+---------------------------+
|  Channel 1 : Tension (V)  |
+---------------------------+

Filtering Data (N=51) : [=====================================================================] 100%

Creating Figures...
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm_Ch1_line=false.png'
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm_Ch1_line=true.png'
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm_Ch1_line=false_filtred_N=51.png'
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm_Ch1_line=true_filtred_N=51.png'

Delta Time Info :
+-------------------------+-------------+
|         Channel         | Tension (V) |
|       Δt Raw (ms)       |    71.10    |
|     Δt Filtred (ms)     |    72.34    |
|    Difference Δt (ms)   |     1.24    |
|                         |             |
|   Time start Raw (ms)   |  -0.071100  |
|    Time End Raw (ms)    |   0.000000  |
|                         |             |
| Time start Filtred (ms) |  -0.072332  |
|  Time End Filtred (ms)  |   0.000010  |
+-------------------------+-------------+

----------------------------------------------------------------------------------------------------

+---------------------------+
|  Channel 2 : Current (A)  |
+---------------------------+

Filtering Data (N=51) : [=====================================================================] 100%

Creating Figures...
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm_Ch2_line=false.png'
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm_Ch2_line=true.png'
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm_Ch2_line=false_filtred_N=51.png'
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm_Ch2_line=true_filtred_N=51.png'

Delta Time Info :
+-------------------------+-------------+
|         Channel         | Current (A) |
|       Δt Raw (ms)       |    71.09    |
|     Δt Filtred (ms)     |    72.30    |
|    Difference Δt (ms)   |     1.21    |
|                         |             |
|   Time start Raw (ms)   |  -0.071100  |
|    Time End Raw (ms)    |  -0.000006  |
|                         |             |
| Time start Filtred (ms) |  -0.072310  |
|  Time End Filtred (ms)  |  -0.000014  |
+-------------------------+-------------+

Creating Combined Excel Sheet :

inserting Data ... OK
inserting Charts ... OK
Closing Excel File ... OK
Excel file created with success : 20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm.xlsx


========================================   PROCESSING FINISHED   ========================================


```
