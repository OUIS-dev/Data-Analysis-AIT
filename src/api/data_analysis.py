'''
---------------------------------------
|   Author :  Mohamed El Amine OUIS   |
---------------------------------------

Description     :   Library to analyse data from 'Oscilator' CSV files containigs
                    Tension and Current Data for specific time period.
                    Each channel has it's own CSV file :
                        channel 1 : Tension (V) [..._Ch1.csv]
                        channel 2 : Current (A) [..._Ch2.csv]
                    
                    This library has a CSV_Object Class.
                    And in the same CSV_object class there is a child class 'Delta'
                    [delta_time, idx_start, time_start, value_start, idx_end, time_end, 
                    value_end]

Release Date    : 08-08-2020

Actual Version  : 1.1

Changelog:

v1.1:
    - Multiple changes to enhence console display info

Changelog:
v1.0:
    - Added CSV_Object Class and Delta child Class in it.
    - Added Delta Class to store delta time info.
    - CSV_Object :multiple methods implemented in it :
        - _extract_info_from_file_name()
        - _load_csv_data()
        - generate_excel()
        - create_plot_data()
        - plot_show()
        - apply_moving_average()
        - calculate_delta_time()
            - _extract_indexes()
    
    - Added Functions :
        - enter_directory()
        - create_dir(directory, dir_name:str='Processed')
        - input_combined()
        - get_file_names_list(directory:str, extension:str='csv', combined:bool=False)
            - _select_files_function(file_names)
            - _get_multiple_opt1(str_input, file_names)
            - _get_multiple_opt2(str_input, file_names)
        - generate_excel_combined(obj_ch1, obj_ch2)
            - _write_header_cell_sheet_data(worksheet, cell_format_header)
            - _write_header_cell_sheet_charts(worksheet_charts, cell_format_header)
            - _write_delta_time_info_to_sheet_charts(worksheet_charts, obj, row:int, cell_format_voltage, cell_format_time)
        - add_chart_to_workbook(workbook, obj_ch1, obj_ch2, filtred:bool=False, width:int=0, height:int=450)
        - print_dict_PrettyTable()
        - print_processing_info(CSV_Object)

'''


import os, csv, xlsxwriter
import matplotlib.pyplot as plt
from datetime import datetime
import statistics, progressbar
from colorama import Fore, Back, Style, init
init(autoreset=True)
from prettytable import PrettyTable

# CSV CLASS : ################################################################################

class Delta():
    delta_time = None
    idx_start = None
    time_start = None
    value_start = None
    idx_end = None
    time_end = None
    value_end = None
        
class CSV_Object():
    def __init__(self, file_path, processed_dir:str='', delimiter:str=',', calibration_value:float=1):
        self.file_path = file_path
        self.file_name = os.path.basename(self.file_path)
        self.directory = os.path.dirname(self.file_path)
        self.processed_dir = processed_dir
        self.delimiter = delimiter
        self.calibration_value = calibration_value
        self.data = []      # List of 'dict' type elements --> data = [{'time':time, 'value':value}, ...]
        self.times = []     # List of 'float' type elements
        self.values = []    # List of 'float' type elements

        self.values_filtred = []  # used to store the filtred values
        self.threshold_low = None
        self.threshold_high = None
        self.delta_time = None

        self.aquisition_date = None
        self.channel_number = None

        self.record_length = None
        self.sample_interval = None
        self.trigger_point = None
        self.trigger_time = None
        self.tension = None
        self.horizontal_offset = None

        self.Delta = Delta()
        self.DeltaFiltred = Delta()

        self._extract_info_from_file_name()
        self._load_csv_data()

    def _extract_info_from_file_name(self):
        # Extract Aquisition Date :
        aquisition_date_str = self.file_name[:8]
        self.aquisition_date = datetime.strptime(aquisition_date_str, '%Y%m%d')

        # Extract Channel Number :
        self.channel_number = int(self.file_name[-5:-4])


    def _load_csv_data(self):
        with open(self.file_path, newline = '') as csvfile:
            csv_reader = csv.reader(csvfile, delimiter = self.delimiter)

            # Get File Info :
            for i, row in enumerate(csv_reader):
                time    = float(row[3])
                value   = float(row[4])
                
                if self.calibration_value != 1.0:
                    value = value * self.calibration_value

                self.data.append({
                                'time':time,
                                'value':value
                            })
                self.times.append(time)
                self.values.append(value)

                if i == 0:
                    try:
                        self.record_length = int(row[1])     # Points
                    except Exception as e:
                        print(f"{Fore.RED}Could not convert 'record_length' !\nError Message : {e}")

                if i == 1:
                    try:
                        self.sample_interval = float(row[1])     # s
                    except Exception as e:
                        print(f"{Fore.RED}Could not convert 'sample_interval' !\nError Message : {e}")

                if i == 2:
                    try:
                        self.trigger_point = int(row[1])     # Samples
                    except Exception as e:
                        print(f"{Fore.RED}Could not convert 'trigger_point' !\nError Message : {e}")

                if i == 3:
                    try:
                        self.trigger_time = float(row[1])     # s
                    except Exception as e:
                        print(f"{Fore.RED}Could not convert 'trigger_time' !\nError Message : {e}")

                if i == 4:
                    try:
                        pass
                        #self.tension = float(row[1])
                    except Exception as e:
                        print(f"{Fore.RED}Could not convert 'tension' !\nError Message : {e}")
                
                if i == 5:
                    try:
                        self.horizontal_offset = float(row[1])   # s
                    except Exception as e:
                        print(f"{Fore.RED}Could not convert 'horizontal_offset' !\nError Message : {e}")

                if i == 6:
                    break  

            # Extract Data :        
            for i, row in enumerate(csv_reader):
                time    = float(row[3])
                value   = float(row[4])
                if self.calibration_value != 1.0:
                    value = value * self.calibration_value
                self.data.append({
                                'time':time,
                                'value':value
                            })
                
                self.times.append(time)
                self.values.append(value)


    def generate_excel(self):
        print('\nCreating Excel Sheet ...', end=' ')
        excel_file_name = os.path.splitext(self.file_name)[0] + '.xlsx'
        excel_file_path = os.path.join(self.directory, self.processed_dir, excel_file_name)

        try:
            workbook = xlsxwriter.Workbook(excel_file_path)
            worksheet = workbook.add_worksheet()

            # Excel header :
            cell_format = workbook.add_format({'bold': True, 'font_color': 'white', 'align':'center', 'bg_color':'#000033'})

            worksheet.write(0, 0, 'Time (s)', cell_format)
            if self.channel_number == 1:
                worksheet.write(0, 1, 'Tension (V)', cell_format)
                worksheet.write(0, 2, 'Tension (V) Filtred', cell_format)
                cell_num_format = workbook.add_format({'num_format' : '0.00', 'align':'center'})

            else:
                worksheet.write(0, 1, 'Current (A)', cell_format)
                worksheet.write(0, 2, 'Current (A) Filtred', cell_format)
                cell_num_format = workbook.add_format({'num_format' : '0.00000000000', 'align':'center'})
            
            # Insert data to Excel file :
            row = 1
            # Insert data to Excel file :
            cell_format_time = workbook.add_format({'num_format' : '0.000000', 'align':'center'})
        
            for i in range(self.record_length):
                worksheet.write_number(row, 0, self.times[i], cell_format_time)
                worksheet.write_number(row, 1, self.values[i], cell_num_format)
                worksheet.write_number(row, 2, self.values_filtred[i], cell_num_format)
                row +=1

            worksheet.set_column(0, 0, 15)
            worksheet.set_column(1, 1, 20)
            worksheet.set_column(2, 2, 22)
            worksheet.autofilter(0, 0, self.record_length, 2)
            workbook.close()    # Close the file
            
            print(f'{Fore.GREEN}OK')
            print(f"Excel file created with success : {excel_file_name}")

        except Exception as e:
            print(f"{Fore.RED}Could not create Excel file !! \nFile path : {excel_file_path}\nError Message : {e}")

            
    def create_plot_data(self, x_label:str='Time (s)', y_label:str='', filtred:bool=False,
                         title:str='', file_name:str='', dpi:int=300, add_vline:bool=True,
                         grid:bool=True, black_bg:bool=True):
        
        x_axis_data = self.times
        if filtred:
            y_axis_data = self.values_filtred
            Delta = self.DeltaFiltred
        else:
            y_axis_data = self.values
            Delta = self.Delta


        w, h = plt.figaspect(0.5)
        fig = plt.figure(figsize=(w, h))
        
        if title == '':
            fig.suptitle(self.file_name, fontsize=14, fontweight='bold')
        else:
            fig.suptitle(title, fontsize=14, fontweight='bold')
        
        ax = fig.add_subplot(111)
                
        ax.set_xlabel(x_label)

        if self.channel_number == 1:
            ax.plot(x_axis_data, y_axis_data, linewidth=1, color='yellow')
            time_v_pos = 1 * self.calibration_value
            delta_v_pos = 31 * self.calibration_value
            if y_label == '':
                ax.set_ylabel('Tension (V)')
        else:
            ax.plot(x_axis_data, y_axis_data, linewidth=1, color='cyan')
            time_v_pos = 0.08 * self.calibration_value
            delta_v_pos = 0.70 * self.calibration_value
            if y_label == '':
                ax.set_ylabel('Current (A)')
        
        if black_bg:
            ax.set_facecolor((0,0,0))
            text_color = 'white'
        else:
            text_color = 'black'

        if add_vline:
            ax.axvline(x=Delta.time_start, linewidth=2)
            ax.text(Delta.time_start + 0.003, time_v_pos, f'Time Start\n{round(Delta.time_start * 1000, 2)} ms', color=text_color)

            ax.axvline(x=Delta.time_end, linewidth=2)
            ax.text(Delta.time_end + 0.003, time_v_pos, f'Time End\n{round(Delta.time_end * 1000, 2)} ms', color=text_color)

            h_pos = float(Delta.time_end + Delta.time_start)/2 - 0.02
            ax.text(h_pos, delta_v_pos, rf'$Δt = {Delta.delta_time}$ ms', fontsize=12, color=text_color) 
        
        if grid:
            ax.grid()
        
        if file_name == '':
            image_file = os.path.splitext(self.file_name)[0] + '.png'
        else:
            image_file = file_name
        
        image_path = os.path.join(self.directory, self.processed_dir, image_file)
        
        try:
            plt.savefig(image_path, dpi=dpi)
            print(f"Figure 'png' file created with success : '{image_file}'")
        except Exception as e:
            print(f"{Fore.RED}\nCould not create Figure 'png' file !! \nPNG path : {image_path}\nError Message : {e}")



    def plot_show(self):
        plt.show()



    def apply_moving_average(self, N:int=31, filter_step:int=1):
        self.values_filtred = []
        n = int(N/2)
        self.values_filtred [:n-1] = self.values[:n-1]
        print('')
        bar = progressbar.ProgressBar(maxval=self.record_length-N+1, widgets=[progressbar.Bar('=', f'Filtering Data (N={N}) : [', ']'), ' ', progressbar.Percentage()], term_width=100)
        bar.start()
        idx = 0
        for i in range(n, self.record_length - n + 1, filter_step):
            mean_value = statistics.mean(self.values[i-n:i+n+1])
            self.values_filtred.append(mean_value)
            bar.update(idx)
            idx += 1
        bar.finish()
        self.values_filtred[self.record_length-n:self.record_length] = self.values[self.record_length-n:]
    


    def _extract_indexes(self, values):
        if self.channel_number == 1:
            for i, value in enumerate(values):
                if float(value) < self.threshold_low :
                    time_start_idx = i

            for i, value in enumerate(values[::-1]):
                if float(value) > self.threshold_high :
                    time_end_idx = self.record_length - i

        else:
            for i, value in enumerate(values):
                if float(value) > self.threshold_high :
                    time_start_idx = i
            
            for i, value in enumerate(values[::-1]):
                if float(value) < self.threshold_low :
                    time_end_idx = self.record_length - i
        
        return [time_start_idx, time_end_idx]

    def calculate_delta_time(self, threshold_low=None, threshold_high=None):
        time_start_idx = None
        time_end_idx = None
        
        if threshold_low is None:
            if self.channel_number == 1:    # Tension (V)
                self.threshold_low = float(3) * self.calibration_value
                self.threshold_high = float(30) * self.calibration_value
            else:                           # Current (A)
                self.threshold_low = float(0.1) * self.calibration_value
                self.threshold_high = float(0.7) * self.calibration_value
        
        # RAW DATA
        time_start_idx, time_end_idx = self._extract_indexes(self.values)   # Detect Start/End of signal variation
        delta_time = round( abs(self.times[time_end_idx] - self.times[time_start_idx]) * 1000, 2)
        self.Delta.delta_time = delta_time
        self.Delta.idx_start = time_start_idx
        self.Delta.idx_end = time_end_idx
        self.Delta.time_start = self.times[time_start_idx]
        self.Delta.time_end = self.times[time_end_idx]
        self.Delta.value_start = self.values[time_start_idx]
        self.Delta.value_end = self.values[time_end_idx]

        # FILTRED DATA
        time_start_idx_filtred, time_end_idx_filtred = self._extract_indexes(self.values_filtred)   # Detect Start/End of signal variation
        delta_time_filtred = round( abs(self.times[time_end_idx_filtred] - self.times[time_start_idx_filtred]) * 1000, 2)
        self.DeltaFiltred.delta_time = delta_time_filtred
        self.DeltaFiltred.idx_start = time_start_idx_filtred
        self.DeltaFiltred.idx_end = time_end_idx_filtred
        self.DeltaFiltred.time_start = self.times[time_start_idx_filtred]
        self.DeltaFiltred.time_end = self.times[time_end_idx_filtred]
        self.DeltaFiltred.value_start = self.values[time_start_idx_filtred]
        self.DeltaFiltred.value_end = self.values[time_end_idx_filtred]




# FUNCTIONS :   ############################################################################

def enter_directory():
    while True:
        print("\nScript Started ...")
        directory = input("\nEnter 'CSV' Directory : ")
        #directory = "/home/operator/Documents/dev/data_analysis_ait/data_samples/"
        if not os.path.isdir(directory):
            print(f"{Fore.RED}Please Enter a correct directory !!")
        else:
            break
    return directory


def create_dir(directory, dir_name:str='Processed'):
    try:
        print(f"\nCreating New Directory : '{dir_name}' in the working directory\n'{directory}' ...")
        dir_path = os.path.join(directory, dir_name)
        os.mkdir(dir_path)
    except OSError:
        if os.path.isdir(dir_path):
            print(f"Directory aleady exists : {dir_path}")
            return True
        else:
            print (f"{Fore.RED}Creation of the directory failed ! \nDirectory path : {dir_path}")
            return False
    else:
        print (f"{Fore.GREEN}Successfully created the directory. \nDirectory path : {dir_path}")
        return True


def input_combined():
    while True:
        print(f"\n{'-'*44}   Combined Option   {'-'*45}")
        str_input = input('''\nProcess the two channels combined [yes or 1]
Process Each channel file one by one [no, 0]
Please select an option : ''')
        if str_input.lower() in ['true', 'yes', 't', 'y', '1']:
            return True
        elif str_input.lower() in ['false', 'no', 'f', 'n', '0']:
            return False
        else:
            print(f"{Fore.RED}Please input a correct value !!!")


def get_file_names_list(directory:str, extension:str='csv', combined:bool=False):
    file_names = [f for f in os.listdir(os.path.join(directory)) if f.endswith('.' + extension)]
    file_names.sort()
    
    while True:
        # print all files in the folder :
        print(f"\n\n{'-'*47}   Files List   {'-'*47}\n")
        print(f"Directory : {directory}\n")
        
        if combined:
            file_names_combined = []
            for file_name in file_names:
                if file_name [-7:] == 'Ch1.csv':
                    file_name_ch2 = file_name[:-7] + 'Ch2.csv'
                    # check if file_name_ch2 exists :
                    ret_val = os.path.isfile(os.path.join(directory, file_name_ch2))
                    if ret_val:
                        file_names_combined.append((file_name, file_name_ch2))  # append the list with a tuple (file_name_ch1, file_name_ch2)
                    else:
                        print(f"{Fore.RED}This File '{file_name}' has No 'Ch2' file !! ")
            
            for i, file_name in enumerate(file_names_combined):
                print(f"index : [{i}] --- File Name : {file_name[0][:-8]}")
            
            file_names_list = _select_files_function(file_names_combined)
            
            if file_names_list:
                return file_names_list

        else:

            for i, file_name in enumerate(file_names):
                print(f"index : [{i}] --- File Name : {file_name}")
            
            file_names_list = _select_files_function(file_names)
            
            if file_names_list:
                return file_names_list


def _select_files_function(file_names):
    print(f"\n\n{'-'*44}   Index Selection   {'-'*45}\n")
    str_input = input(f'''Info :
One File : input the file index
Multiple Files (option 1) : input files indexes and separate between them with '.' or ',' or ';'
Multiple Files (option 2) : input idx_start and idx_end separated by ':' to get all files having indexes between them
All File : just input 'all' or 'a'

input index(es) : ''')


    if str_input.lower() in ['all', 'a']:
        print("All files are selected.")
        return file_names
    
    else:
        if str_input.isnumeric():
            if int(str_input) < len(file_names):
                file_name = file_names[int(str_input)]
                print(f"\n\n{'-'*45}   Selected file   {'-'*46}\n")
                print(f"index : [{int(str_input)}] --- File Name : {file_name}\n")
                return [file_name] # return the file_name in the list
            else:
                print(f"{Fore.RED}Index is not in the list! please select a correct index.")
                return []
        
        else:
            file_names_list = _get_multiple_opt1(str_input, file_names)
            if file_names_list:
                return file_names_list
            else:
                file_names_list = _get_multiple_opt2(str_input, file_names)
                if file_names_list:
                    return file_names_list
                else:
                    print(f"{Fore.RED}Please input values correctly !!!")


def _get_multiple_opt1(str_input, file_names):
    file_names_list = []
    for separator in [',', ';', '.']:
        f = str_input.find(separator)
        if f > -1 : # if the element in the str it will return positif value (0 included)
                    # if the value equal -1 :  the element serached is not found.
            
            index_list = str_input.split(separator)

            if index_list:
                #print("\nSelected files :\n----------------")
                print(f"\n\n{'-'*45}   Selected Files   {'-'*45}\n")
                for i in index_list:
                    i = int(i)
                    if i >= len(file_names):
                        print(f"{Fore.RED}Index : [{i}] --- This index is not in the list !! ")
                        continue
                    else:
                        file_names_list.append(file_names[i])
                        print(f"index : [{i}] --- File Name : {file_names[i]}")
    
    return file_names_list


def _get_multiple_opt2(str_input, file_names):

    file_names_list = []
    index_list = str_input.split(':')
    if index_list:
        try:
            index_start = int(index_list[0])
            index_end = int(index_list[1]) + 1
            file_names_list = file_names[index_start : index_end]
            
            print(f"\n\n{'-'*45}   Selected Files   {'-'*45}\n")
            for i, file_name in enumerate(file_names_list):
                print(f"index : [{i+index_start}] --- File Name : {file_name}")
                
        except Exception as e:
            print(f"{Fore.RED} Error Message : {e}")

        return file_names_list



def _write_header_cell_sheet_data(worksheet, cell_format_header):
    worksheet.write(0, 0, 'Time', cell_format_header)
    worksheet.write(0, 1, 'Tension (V)', cell_format_header)
    worksheet.write(0, 2, 'Tension (V) Filtered', cell_format_header)
    worksheet.write(0, 3, 'Current (A)', cell_format_header)
    worksheet.write(0, 4, 'Current (A) Filtered', cell_format_header)
    
    worksheet.set_column(0, 0, 15)
    worksheet.set_column(1, 1, 15)
    worksheet.set_column(2, 2, 22)
    worksheet.set_column(3, 3, 20)
    worksheet.set_column(4, 4, 22)

def _write_header_cell_sheet_charts(worksheet_charts, cell_format_header):
    worksheet_charts.write(0, 1, 'Δt (ms)', cell_format_header)
    worksheet_charts.write(0, 2, 'Time start (ms)', cell_format_header)
    worksheet_charts.write(0, 3, 'Time start Index', cell_format_header)
    worksheet_charts.write(0, 4, 'Time End (ms)', cell_format_header)
    worksheet_charts.write(0, 5, 'Time End Index', cell_format_header)
    worksheet_charts.write(1, 0, 'Tension (Raw)', cell_format_header)
    worksheet_charts.write(2, 0, 'Tension (Filtred)', cell_format_header)
    worksheet_charts.write(3, 0, 'Current (Raw)', cell_format_header)
    worksheet_charts.write(4, 0, 'Current (Filtred)', cell_format_header)

    worksheet_charts.set_column("$A1:$F1", 15)

def _write_delta_time_info_to_sheet_charts(worksheet_charts, obj, row:int, cell_format_voltage, cell_format_time):
    worksheet_charts.write_number(row, 1, obj.Delta.delta_time, cell_format_voltage)   # cell voltage has 2 decimal digits 
    worksheet_charts.write_number(row, 2, obj.Delta.time_start, cell_format_time)
    worksheet_charts.write_number(row, 3, obj.Delta.idx_start)
    worksheet_charts.write_number(row, 4, obj.Delta.time_end, cell_format_time)
    worksheet_charts.write_number(row, 5, obj.Delta.idx_end)

    worksheet_charts.write_number(row+1, 1, obj.DeltaFiltred.delta_time, cell_format_voltage)   # cell voltage has 2 decimal digits 
    worksheet_charts.write_number(row+1, 2, obj.DeltaFiltred.time_start, cell_format_time)
    worksheet_charts.write_number(row+1, 3, obj.DeltaFiltred.idx_start)
    worksheet_charts.write_number(row+1, 4, obj.DeltaFiltred.time_end, cell_format_time)
    worksheet_charts.write_number(row+1, 5, obj.DeltaFiltred.idx_end)

def generate_excel_combined(obj_ch1, obj_ch2):
    print('\nCreating Combined Excel Sheet :\n')
    excel_file_name = obj_ch1.file_name[:-8] + '.xlsx'
    excel_file_path = os.path.join(obj_ch1.directory, obj_ch1.processed_dir, excel_file_name)
    

    workbook = xlsxwriter.Workbook(excel_file_path)
    worksheet = workbook.add_worksheet('Data')
    worksheet_charts = workbook.add_worksheet('Charts')

    # Excel cell format :
    cell_format_header = workbook.add_format({'bold': True, 'font_color': 'white', 'align':'center', 'bg_color':'#000033'})
    cell_format_time = workbook.add_format({'num_format' : '0.000000', 'align':'center'})
    cell_format_voltage = workbook.add_format({'num_format' : '0.00', 'align':'center'})
    cell_format_current = workbook.add_format({'num_format' : '0.00000000000', 'align':'center'})

    _write_header_cell_sheet_data(worksheet, cell_format_header)
    _write_header_cell_sheet_charts(worksheet_charts, cell_format_header)
    
    # Insert data to Worksheet Data :
    print('inserting Data ...',end=' ')
    for i in range(obj_ch1.record_length):
        worksheet.write_number(i+1, 0, obj_ch1.times[i], cell_format_time)
        worksheet.write_number(i+1, 1, obj_ch1.values[i], cell_format_voltage)
        worksheet.write_number(i+1, 2, obj_ch1.values_filtred[i], cell_format_voltage)
        worksheet.write_number(i+1, 3, obj_ch2.values[i], cell_format_current)
        worksheet.write_number(i+1, 4, obj_ch2.values_filtred[i], cell_format_current)

    print(f'{Fore.GREEN}OK')
    worksheet.autofilter(0, 0, obj_ch1.record_length, 4)

    # Insert data and charts to Worksheet Charts :
    _write_delta_time_info_to_sheet_charts(worksheet_charts, obj_ch1, 1, cell_format_voltage, cell_format_time)
    _write_delta_time_info_to_sheet_charts(worksheet_charts, obj_ch2, 3, cell_format_voltage, cell_format_time)

    # Add Chart :
    print('inserting Charts ...',end=' ')
    chart_raw = add_chart_to_workbook(workbook, obj_ch1, obj_ch2, filtred=False)
    worksheet_charts.insert_chart('B7', chart_raw)

    chart_filtred = add_chart_to_workbook(workbook, obj_ch1, obj_ch2, filtred=True)
    worksheet_charts.insert_chart('B32', chart_filtred)
    print(f'{Fore.GREEN}OK')
    print('Closing Excel File ...',end=' ')
    workbook.close()    # Close the file
    print(f'{Fore.GREEN}OK')
    print(f"Excel file created with success : {excel_file_name}")



def add_chart_to_workbook(workbook, obj_ch1, obj_ch2, filtred:bool=False, width:int=0, height:int=450):
    excel_file_name = obj_ch1.file_name[:-8] + '.xlsx'

     # Create Tension + Current chart :
    chart = workbook.add_chart({'type' : 'line'})
    chart_current = workbook.add_chart({'type': 'line'})
    
    if filtred:
        v_col = 'C'
        c_col = 'E'
        v_name = 'Tension (V) Filtred'
        c_name = 'Current (A) Filtred'
        title = f'{excel_file_name[:-5]}_Filtred'
        if width == 0:
            width = 900
    else:
        v_col = 'B'
        c_col = 'D'
        v_name = 'Tension (V)'
        c_name = 'Current (A)'
        title = f'{excel_file_name[:-5]}_Raw'
        if width == 0:
            width = 860
    
    chart.add_series({
        'name': f'=Data!${v_col}$1',
        'categories': f'=Data!$A$1:$A${obj_ch1.record_length}',
        'values': f'=Data!${v_col}$1:${v_col}${obj_ch1.record_length}',
        'line': {'color': 'yellow'},
        'name_font': {'color':'white'},
        'num_font': {'color':'white'}
        })

    chart_current.add_series({
        'name': f'=Data!${c_col}$1',
        'categories': f'=Data!$A$1:$A${obj_ch1.record_length}',
        'values': f'=Data!${c_col}$1:${c_col}${obj_ch1.record_length}',
        'y2_axis': True,
        'line': {'color': 'cyan'},
        'name_font': {'color':'white'},
        'num_font': {'color':'white'}
        })

    # Combine charts together : 
    chart.combine(chart_current)
    
    chart.set_title({'name': title, 'name_font': {'color':'white'}, 'num_font': {'color':'white'}})
    chart.set_x_axis({
        'name': 'Time (s)',
        'name_font': {'color':'white'},
        'num_font': {'color':'white', 'rotation': -90},
        #'major_gridlines': {'visible': True},
        })
    
    chart.set_y_axis({
        'name': v_name,
        'name_font': {'color':'white'},
        'num_font': {'color':'white'},
        })
    
    chart_current.set_y2_axis({
        'name': c_name,
        'name_font': {'color':'white'},
        'num_font': {'color':'white'}
        })
    
    chart.set_size({'width': width, 'height': height})
    chart.set_chartarea({'fill':   {'color': 'black'}})
    chart.set_plotarea({'fill':   {'color': 'black'}})
    chart.set_legend({'font': {'color':'white'}})
    return chart
    

def print_dict_PrettyTable(dict:dict, orientation:str='vertical', header:bool=True, padding_width:int=0, key='key', key_allign='', value='value', value_allign='', hrules='ALL', vrules='ALL'):
    ''' Dependencies : from prettytable import PrettyTable  (pip install)
    :param dict: dictionary to be printed in PrettyTable format
    :param orientation: which orientation to print the Pretty Table ('v', 'vertical', 'h' or 'horizontal')
    :param header : flag to print or not the table header.
    :param padding_width : padding between columns of the table.
    '''
    table = PrettyTable(header=header, padding_width=padding_width)

    if orientation == 'v' or orientation == 'vertical':
        table.field_names = dict.keys()
        table.add_row(dict.values())

    elif orientation == 'h' or orientation == 'horizontal':
        table.add_column(key,[*dict.keys()])
        table.align[key] = key_allign

        table.add_column(value,[*dict.values()])
        table.align[value] = value_allign

    else:
        print(f"{Fore.RED}Please enter a correct value of orientation ('v', 'vertical', 'h' or 'horizontal') !")
        return

    print(table)

    

def print_processing_info(obj):
    if obj.channel_number == 1:
        channel_name = 'Tension (V)'
    else:
        channel_name = 'Current (A)'

    info_to_print = {
        'Channel' : channel_name,
        'Δt Raw (ms)' : f'{obj.Delta.delta_time:.2f}',
        'Δt Filtred (ms)' : f'{obj.DeltaFiltred.delta_time:.2f}',
        'Difference Δt (ms)' : f'{abs(obj.Delta.delta_time - obj.DeltaFiltred.delta_time):.2f}',
        '':'',
        'Time start Raw (ms)' : f'{obj.Delta.time_start:.6f}',
        'Time End Raw (ms)' : f'{obj.Delta.time_end:.6f}',
        ' ':'',
        'Time start Filtred (ms)' : f'{obj.DeltaFiltred.time_start:.6f}',
        'Time End Filtred (ms)' : f'{obj.DeltaFiltred.time_end:.6f}',
    }
    print("\nDelta Time Info :")
    print_dict_PrettyTable(info_to_print, orientation='h', header=False)
