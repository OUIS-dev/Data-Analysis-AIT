'''
---------------------------------------
|   Author :  Mohamed El Amine OUIS   |
---------------------------------------

Description     :   Main Code to analyse data from 'Oscilator' CSV files containigs
                    Tension and Current Data for specific time period.
                    Each channel has it's own CSV file :
                        channel 1 : Tension (V) [..._Ch1.csv]
                        channel 2 : Current (A) [..._Ch2.csv]
                    
                    The output files are : 
                        - Excel Sheet containing all data and charts + info
                        - Figures images for each channel with Delta time (ms)

Release Date    : 08-08-2020

Actual Version  : 1.1

Changelog:

v1.1:
    - Added 'config.ini' file to set all default parameters inside.
    - Enhenced Console Display.
    - Fixed file name in figure title.
    - Other minor changes.


v1.0:
    - Select Working csv_directory where the CSV files are stored
    - Creation of new csv_directory to save the processed data : Excel sheets + images
    - The script lists all 'CSV' files in the working directory, each one with its own index
    - The user select which indexes to process. The selection is possible using these methods:
        - One File : input the file index
        - Multiple Files (option 1) : input files indexes and separate wetween them with '.' or ',' or ';'
        - Multiple Files (option 2) : input idx_start and idx_end separated by ':' to get all files having indexes between them
        - All File : just input 'all' or 'a'
    - It is possible to process each file seprately, or as a combined channels (Ch1_Tension + Ch2_Current)
    - The implemented filter used to filter the data is SMA : Simple Moving Average with a windows Size = N (fixed in config.ini)
    
    - The process of combined files ch1 + ch2:
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

    - The process of all files separetly :
        - Select files indexes to process
        - For each File :
            - Create new folder to save Excel + images
            - Start by creating a CSV_Object instance to load the data from CSV file
            - apply_moving_average()
            - calculate_delta_time() for both raw and filtred data
            - create_plot_data() for both raw and filtred data
            - Generate exel file with graphs using this method : generate_excel()

'''

import os, xlsxwriter, statistics, progressbar
from colorama import Fore, Back, Style, init
from api import data_analysis
init(autoreset=True)


###    Functions    ############################################################################

def combined_files_processing(csv_directory, file_names_list):
    for i, file_compined in enumerate(file_names_list):
        file_name_ch1 = file_compined[0]
        file_name_ch2 = file_compined[1]
        
        print(f"{Fore.BLACK}{Back.WHITE}\n\n[{i+1}/{len(file_names_list)}] Processing File : {file_name_ch1[:-8]}\n")

        obj_ch1 = data_analysis.CSV_Object(file_path=os.path.join(csv_directory, file_name_ch1), combined_flag=True)
        obj_ch2 = data_analysis.CSV_Object(file_path=os.path.join(csv_directory, file_name_ch2), combined_flag=True)

        # Channel 1 :   ------------------------------------------
        print('\n' + '-'*100 + '\n')
        print(f'{Fore.BLACK}{Back.WHITE}+---------------------------+')
        print(f'{Fore.BLACK}{Back.WHITE}|  Channel 1 : Tension (V)  |')
        print(f'{Fore.BLACK}{Back.WHITE}+---------------------------+')

        obj_ch1.apply_moving_average()  # it will generate data into 'self.values_filtred' 
        obj_ch1.calculate_delta_time()

        # Create plot images :
        # RAW
        print("\nCreating Figures...")
        file_name=f'{obj_ch1.file_name[:-4]}_line=false.png'
        obj_ch1.create_plot_data(filtred=False, add_vline=False, file_name=file_name, title=file_name)
        
        file_name=f'{obj_ch1.file_name[:-4]}_line=true.png'
        obj_ch1.create_plot_data(filtred=False, add_vline=True, file_name=file_name, title=file_name)
        
        # FILTRED
        file_name=f'{obj_ch1.file_name[:-4]}_line=false_filtred_N={obj_ch1.N}.png'
        obj_ch1.create_plot_data(filtred=True, add_vline=False, file_name=file_name, title=file_name)

        file_name=f'{obj_ch1.file_name[:-4]}_line=true_filtred_N={obj_ch1.N}.png'
        obj_ch1.create_plot_data(filtred=True, add_vline=True, file_name=file_name, title=file_name)

        data_analysis.print_processing_info(obj_ch1)

        # Channel 2 :   ------------------------------------------
        print('\n' + '-'*100 + '\n')
        print(f'{Fore.BLACK}{Back.WHITE}+---------------------------+')
        print(f'{Fore.BLACK}{Back.WHITE}|  Channel 2 : Current (A)  |')
        print(f'{Fore.BLACK}{Back.WHITE}+---------------------------+')

        obj_ch2.apply_moving_average()  # it will generate data into 'self.values_filtred' 
        obj_ch2.calculate_delta_time()

        # Create plot images :
        # RAW
        print("\nCreating Figures...")
        file_name=f'{obj_ch2.file_name[:-4]}_line=false.png'
        obj_ch2.create_plot_data(filtred=False, add_vline=False, file_name=file_name, title=file_name)

        file_name=f'{obj_ch2.file_name[:-4]}_line=true.png'
        obj_ch2.create_plot_data(filtred=False, add_vline=True, file_name=file_name, title=file_name)
        
        # FILTRED
        file_name=f'{obj_ch2.file_name[:-4]}_line=false_filtred_N={obj_ch2.N}.png'
        obj_ch2.create_plot_data(filtred=True, add_vline=False, file_name=file_name, title=file_name)

        file_name=f'{obj_ch2.file_name[:-4]}_line=true_filtred_N={obj_ch2.N}.png'
        obj_ch2.create_plot_data(filtred=True, add_vline=True, file_name=file_name, title=file_name)

        data_analysis.print_processing_info(obj_ch2)

        # Generate combine Excel File :
        data_analysis.generate_excel_combined(obj_ch1, obj_ch2)


def separate_files_processing(csv_directory, file_names_list):
    for i, file_name in enumerate(file_names_list):
        print(f"{Fore.BLACK}{Back.WHITE}\n\n[{i+1}/{len(file_names_list)}] Processing File : {file_name}\n")
               
        obj = data_analysis.CSV_Object(file_path=os.path.join(csv_directory, file_name))
        
        obj.apply_moving_average()  # it will generate data into 'self.values_filtred' 
        obj.calculate_delta_time()

        # Create plot images :
        # RAW
        print("\nCreating Figures...")
        file_name=f'{obj.file_name[:-4]}_line=false.png'
        obj.create_plot_data(filtred=False, add_vline=False, file_name=file_name, title=file_name)
        
        file_name=f'{obj.file_name[:-4]}_line=true.png'
        obj.create_plot_data(filtred=False, add_vline=True, file_name=file_name, title=file_name)
        
        # FILTRED 
        file_name=f'{obj.file_name[:-4]}_line=false_filtred_N={obj.N}.png'
        obj.create_plot_data(filtred=True, add_vline=False, file_name=file_name, title=file_name)
        
        file_name=f'{obj.file_name[:-4]}_line=true_filtred_N={obj.N}.png'
        obj.create_plot_data(filtred=True, add_vline=True, file_name=file_name, title=file_name)

        data_analysis.print_processing_info(obj)
        obj.generate_excel()


def main():

    # Enter CSV Working Directory :
    csv_directory = data_analysis.enter_directory()

    combined = data_analysis.input_combined()

    # Get file name list to be processed :
    file_names_list = data_analysis.get_file_names_list(csv_directory, extension='csv', combined=combined)

    # -------   Processing :   -------
    if combined:
        combined_files_processing(csv_directory, file_names_list)
    else:
        separate_files_processing(csv_directory, file_names_list)


if __name__ == '__main__':
    main()
    print(f"{Fore.BLACK}{Back.WHITE}\n\n{'='*40}   PROCESSING FINISHED   {'='*40}\n")