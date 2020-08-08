# Data-Analysis-AIT


## Author :  Mohamed El Amine OUIS


## Description
Main Code to analyse data from 'Oscilator' CSV files containings Tension and Current Data for specific time period.
Each channel has it's own CSV file :
- channel 1 : Tension (V) [..._Ch1.csv]
- channel 2 : Current (A) [..._Ch2.csv]

The output files are :
- Excel Sheet containing all data and charts + info.
- Figures images for each channel with Delta time (ms).

### Release Date    : 08-08-2020

### Actual Version  : 1.1

# Installation

Required Python 3+ to run

Clone repository :
```sh
$ git clone https://github.com/OUIS-dev/Data-Analysis-AIT.git
```

Change directory to the new downloaded directory :
```sh
$ cd Data-Analysis-AIT
```

Create new virtual envirement : 
```sh
$ virtualenv env

```
Activate the virtualenv :
- Linux/Mac
```sh
$ source env/bin/activate
```
- Windows
```sh
$ \env\Scripts\activate.bat
```

install required python packages:
```sh
(env) $ pip install -r requirement.txt
```

Run the script :
```sh
(env) $ cd src
(env) $ python main.py
```

# Usage
```sh
(env) [ouis@ouis-pc src]$ python main.py

Script Started ...

Enter 'CSV' Directory : data_samples

Creating New Directory : 'Processed Data' in the working directory
'data_samples' ...
Successfully created the directory. 
Directory path : data_samples/Processed Data                                                                           

--------------------------------------------   Combined Option   ---------------------------------------------

Process the two channels combined [yes or 1]
Process Each channel file one by one [no, 0]
Please select an option : 1


-----------------------------------------------   Files List   -----------------------------------------------

Directory : data_samples

index : [0] --- File Name : 20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm
index : [1] --- File Name : 20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm


--------------------------------------------   Index Selection   ---------------------------------------------

Info :
One File : input the file index
Multiple Files (option 1) : input files indexes and separate between them with '.' or ',' or ';'
Multiple Files (option 2) : input idx_start and idx_end separated by ':' to get all files having indexes between them
All File : just input 'all' or 'a'

input index(es) : 0.1


---------------------------------------------   Selected Files   ---------------------------------------------

index : [0] --- File Name : ('20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm_Ch1.csv', '20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm_Ch2.csv')
index : [1] --- File Name : ('20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm_Ch1.csv', '20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm_Ch2.csv')


[1/2] Processing File : 20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm                                                   


Creating New Directory : '20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm' in the working directory
'data_samples/Processed Data' ...
Successfully created the directory. 
Directory path : data_samples/Processed Data/20200729__LTN R5_8A_32V_256W_20N_test_OP1Wfm                              

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
|       Δt Raw (ms)       |    71.09    |
|     Δt Filtred (ms)     |    72.32    |
|    Difference Δt (ms)   |     1.23    |
|                         |             |
|   Time start Raw (ms)   |  -0.071100  |
|    Time End Raw (ms)    |  -0.000006  |
|                         |             |
| Time start Filtred (ms) |  -0.072332  |
|  Time End Filtred (ms)  |  -0.000016  |
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


[2/2] Processing File : 20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm                                                   


Creating New Directory : '20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm' in the working directory
'data_samples/Processed Data' ...
Successfully created the directory. 
Directory path : data_samples/Processed Data/20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm                              

----------------------------------------------------------------------------------------------------

+---------------------------+
|  Channel 1 : Tension (V)  |
+---------------------------+

Filtering Data (N=51) : [=====================================================================] 100%

Creating Figures...
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm_Ch1_line=false.png'
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm_Ch1_line=true.png'
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm_Ch1_line=false_filtred_N=51.png'
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm_Ch1_line=true_filtred_N=51.png'

Delta Time Info :
+-------------------------+-------------+
|         Channel         | Tension (V) |
|       Δt Raw (ms)       |    81.89    |
|     Δt Filtred (ms)     |    84.08    |
|    Difference Δt (ms)   |     2.19    |
|                         |             |
|   Time start Raw (ms)   |  -0.081896  |
|    Time End Raw (ms)    |  -0.000006  |
|                         |             |
| Time start Filtred (ms) |  -0.084098  |
|  Time End Filtred (ms)  |  -0.000018  |
+-------------------------+-------------+

----------------------------------------------------------------------------------------------------

+---------------------------+
|  Channel 2 : Current (A)  |
+---------------------------+

Filtering Data (N=51) : [=====================================================================] 100%

Creating Figures...
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm_Ch2_line=false.png'
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm_Ch2_line=true.png'
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm_Ch2_line=false_filtred_N=51.png'
Figure 'png' file created with success : '20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm_Ch2_line=true_filtred_N=51.png'

Delta Time Info :
+-------------------------+-------------+
|         Channel         | Current (A) |
|       Δt Raw (ms)       |    81.89    |
|     Δt Filtred (ms)     |    84.05    |
|    Difference Δt (ms)   |     2.16    |
|                         |             |
|   Time start Raw (ms)   |  -0.081896  |
|    Time End Raw (ms)    |  -0.000006  |
|                         |             |
| Time start Filtred (ms) |  -0.084072  |
|  Time End Filtred (ms)  |  -0.000018  |
+-------------------------+-------------+

Creating Combined Excel Sheet :

inserting Data ... OK
inserting Charts ... OK
Closing Excel File ... OK
Excel file created with success : 20200729__LTN R5_8A_32V_256W_20N_test_OP2Wfm.xlsx


========================================   PROCESSING FINISHED   ========================================


```
