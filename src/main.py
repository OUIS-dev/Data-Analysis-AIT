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
    - Enhenced Console Display.
    - Fixed file name in figure title.
    - Other minor changes.


v1.0:
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


###    Main Function    ############################################################################

def main():

    # Enter CSV Working Directory :
    directory = data_analysis.enter_directory()
    
    # Create new Directory to write in all processed files :
    processed_dir = 'Processed Data'
    data_analysis.create_dir(directory, dir_name=processed_dir)

    combined = data_analysis.input_combined()

    # Get file name list to be processed :
    file_names_list = data_analysis.get_file_names_list(directory, extension='csv', combined=combined)

    N = 51   # filter window : Must be odd number

    # -------   Processing :   -------
    if combined:
        for i, file_compined in enumerate(file_names_list):
            file_name_ch1 = file_compined[0]
            file_name_ch2 = file_compined[1]
            
            print(f"{Fore.BLACK}{Back.WHITE}\n\n[{i+1}/{len(file_names_list)}] Processing File : {file_name_ch1[:-8]}\n")
            output_dir_name = file_name_ch1[:-8]
            data_analysis.create_dir(os.path.join(directory, processed_dir), dir_name=output_dir_name)

            obj_ch1 = data_analysis.CSV_Object(file_path=os.path.join(directory, file_name_ch1),
                                    processed_dir=os.path.join(processed_dir, output_dir_name))
            obj_ch2 = data_analysis.CSV_Object(file_path=os.path.join(directory, file_name_ch2),
                                    processed_dir=os.path.join(processed_dir, output_dir_name),
                                    calibration_value=10)

            # Channel 1 :   ------------------------------------------
            print('\n' + '-'*100 + '\n')
            print(f'{Fore.BLACK}{Back.WHITE}+---------------------------+')
            print(f'{Fore.BLACK}{Back.WHITE}|  Channel 1 : Tension (V)  |')
            print(f'{Fore.BLACK}{Back.WHITE}+---------------------------+')

            obj_ch1.apply_moving_average(N)  # it will generate data into 'self.values_filtred' 
            obj_ch1.calculate_delta_time()

            # Create plot images :
            # RAW
            print("\nCreating Figures...")
            file_name=f'{obj_ch1.file_name[:-4]}_line=false.png'
            obj_ch1.create_plot_data(filtred=False, add_vline=False, file_name=file_name, title=file_name)
            
            file_name=f'{obj_ch1.file_name[:-4]}_line=true.png'
            obj_ch1.create_plot_data(filtred=False, add_vline=True, file_name=file_name, title=file_name)
            
            # FILTRED
            file_name=f'{obj_ch1.file_name[:-4]}_line=false_filtred_N={N}.png'
            obj_ch1.create_plot_data(filtred=True, add_vline=False, file_name=file_name, title=file_name)

            file_name=f'{obj_ch1.file_name[:-4]}_line=true_filtred_N={N}.png'
            obj_ch1.create_plot_data(filtred=True, add_vline=True, file_name=file_name, title=file_name)

            data_analysis.print_processing_info(obj_ch1)

            # Channel 2 :   ------------------------------------------
            print('\n' + '-'*100 + '\n')
            print(f'{Fore.BLACK}{Back.WHITE}+---------------------------+')
            print(f'{Fore.BLACK}{Back.WHITE}|  Channel 2 : Current (A)  |')
            print(f'{Fore.BLACK}{Back.WHITE}+---------------------------+')

            obj_ch2.apply_moving_average(N)  # it will generate data into 'self.values_filtred' 
            obj_ch2.calculate_delta_time()

            # Create plot images :
            # RAW
            print("\nCreating Figures...")
            file_name=f'{obj_ch2.file_name[:-4]}_line=false.png'
            obj_ch2.create_plot_data(filtred=False, add_vline=False, file_name=file_name, title=file_name)

            file_name=f'{obj_ch2.file_name[:-4]}_line=true.png'
            obj_ch2.create_plot_data(filtred=False, add_vline=True, file_name=file_name, title=file_name)
            
            # FILTRED
            file_name=f'{obj_ch2.file_name[:-4]}_line=false_filtred_N={N}.png'
            obj_ch2.create_plot_data(filtred=True, add_vline=False, file_name=file_name, title=file_name)

            file_name=f'{obj_ch2.file_name[:-4]}_line=true_filtred_N={N}.png'
            obj_ch2.create_plot_data(filtred=True, add_vline=True, file_name=file_name, title=file_name)

            data_analysis.print_processing_info(obj_ch2)

            # Generate combine Excel File :
            data_analysis.generate_excel_combined(obj_ch1, obj_ch2)


    else:
        for i, file_name in enumerate(file_names_list):
            print(f"{Fore.BLACK}{Back.WHITE}\n\n[{i+1}/{len(file_names_list)}] Processing File : {file_name}\n")
            #print(f"\n{'*'*120}\n{'*'*120}\n\n[{i+1}/{len(file_names_list)}] Processing File : {file_name}")
            
            output_dir_name = file_name[:-4]
            data_analysis.create_dir(os.path.join(directory, processed_dir), dir_name=output_dir_name)

            if file_name[-5:-4] == '2':
                calibration_value = 10
            else:
                calibration_value = 1
            
            obj = data_analysis.CSV_Object(file_path=os.path.join(directory, file_name),
                                            processed_dir=os.path.join(processed_dir, output_dir_name),
                                            calibration_value=calibration_value)
            
            obj.apply_moving_average(N)  # it will generate data into 'self.values_filtred' 
            obj.calculate_delta_time()

            # Create plot images :
            # RAW
            print("\nCreating Figures...")
            file_name=f'{obj.file_name[:-4]}_line=false.png'
            obj.create_plot_data(filtred=False, add_vline=False, file_name=file_name, title=file_name)
            
            file_name=f'{obj.file_name[:-4]}_line=true.png'
            obj.create_plot_data(filtred=False, add_vline=True, file_name=file_name, title=file_name)
            
            # FILTRED 
            file_name=f'{obj.file_name[:-4]}_line=false_filtred_N={N}.png'
            obj.create_plot_data(filtred=True, add_vline=False, file_name=file_name, title=file_name)
            
            file_name=f'{obj.file_name[:-4]}_line=true_filtred_N={N}.png'
            obj.create_plot_data(filtred=True, add_vline=True, file_name=file_name, title=file_name)

            data_analysis.print_processing_info(obj)
            obj.generate_excel()

        
if __name__ == '__main__':
    main()
    print(f"{Fore.BLACK}{Back.WHITE}\n\n{'='*40}   PROCESSING FINISHED   {'='*40}\n")