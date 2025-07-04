import zipfile
import os
import re
import pandas as pd
import tempfile
import datetime
from pathlib import Path
import json
import math

#---------------------------default settings----------------------------------

folderpath = '../../WAFER SYSTEMS/Tin Lai - Log/30.6.2025/'    
today = datetime.datetime.now()
output_excel_location = today.strftime('xlsx/%d-%m-%Y_error_logs.xlsx')

default_start_date = '2025-06-21'
# default_start_date = False
default_end_date = 'none'
start_date = False
end_date = False
default_error_types = {
    "A1": "检测服务器状态：False",
    "A2": "程序启动时，网络异常，将进入离线模式",
    "B1": "获取工作站状态失败，服务不可访问",
    "B2": "登录页，已到最大重试次数，进入离线模式",
    "boot": "获取到当前的系统默认的代理参数,本地配置为空",
    "language_change": "切换了语言",
    "logout(timeout)": "倒计时结束，返回登录界面",
    "logout(user)": "确认了退出操作"
}
#--------------------------------end of default------------------------------
#--------------------------start of some global var--------------------------
normal_boot_indicator_time = False
#--------------------------end of some global var--------------------------
def parse_log_line(line):
    """Parse a local log line to extract timestamp and message."""
    pattern = r'(\d{2}:\d{2}:\d{2})(\s+\d+\s+|\.\d{3}\s*)\[[^\]]+\](.*)'
    match = re.match(pattern, line.strip())
    if match:
        return {
            'Time': match.group(1),
            'Message': match.group(2).strip()
        }
    return None

def calc_day_sum(year, month, day):
    return year*365 + month*30 + day

def getJSONFILE(path):
    datafile = path
    try:
        if datafile.is_file():    
            with open(datafile, encoding='utf-8') as json_file:
                data = json.load(json_file)
        else:
            data = default_error_types
            with open(datafile, 'w') as json_file:
                json.dump(data, json_file, indent=4)
    except json.JSONDecodeError as e:
        print(f"JSONDecodeError: {e}. Initializing with default data.")
        data = default_error_types
        with open(datafile, 'w') as json_file:
                json.dump(data, json_file, indent=4)
    return data

def get_error_json():
    path = Path('error_types.json')
    return getJSONFILE(path)

def add_error_json(key, content):
    datafile = Path('error_types.json')
    data = get_error_json()
    if key in data:
        return False
    else:
        data[key] = content
        with open(datafile, 'w') as json_file:
            json.dump(data, json_file, indent=4)
        return True

def remove_error_json(key):
    datafile = Path('error_types.json')
    data = get_error_json()
    if key in data:
        del data[key]
        with open(datafile, 'w') as json_file:
            json.dump(data, json_file, indent=4)
        return True
    else:
        return False
        
def get_error_type(message, parsed):
    global normal_boot_indicator_time
    error_types = get_error_json()
    """Determine the error type based on the message content."""
    for error in error_types:
        if re.search(error_types[error], message):
            if error == "language_change" or error == "logout(timeout)" or error == "logout(user)":
                normal_boot_indicator_time = parsed["Time"]
                return False
            if error == "boot":
                if normal_boot_indicator_time:  #case: last indicator exist
                    time_diff = parsed["Time"] - normal_boot_indicator_time
                    if time_diff.total_seconds() > 3*60:    #case: last indicator exists and is > 3 min earlier
                        return "Abnormal boot"
                    else:   #case: last indicator exist and is within 3 min = normal boot
                        return "Normal boot" 
                else:   #case: last indicator dont exist, but is booting
                    return "First boot"
            
            return error
    return False

def process_log_file(filepath, date_str, library, machine):
    """Read a log file and extract error-related data."""
    data = []
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            parsed = parse_log_line(line)  # Use single parsing function
            
               
            if parsed:
                time = datetime.datetime.strptime(parsed['Time'], '%H:%M:%S')   #USING TIME OBJECT for get error type
                parsed.update({"Time": time})
                error_type = get_error_type(line, parsed)  
                if error_type:
                    parsed.update({"Time": time.time()})
                    parsed['Error Type'] = error_type
                    parsed['Date'] = date_str
                    parsed['Library'] = library
                    parsed['Machine'] = machine
                    data.append(parsed)
    return data
# def process_local_log_file(filepath, date_str, library, machine):
#     """Read a local log file and extract error-related data."""
#     data = []
#     with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
#         for line in f:
#             parsed = parse_local_log_line(line)
#             if parsed:
#                 # error_type = get_error_type(parsed['Message'])
#                 error_type = get_error_type(line)
#                 if error_type:
#                     parsed['Error Type'] = error_type
#                     parsed['Date'] = date_str
#                     parsed['Library'] = library
#                     parsed['Machine'] = machine
#                     data.append(parsed)
#     return data


def recursive_walk_for_zip(logs_folder, log_filetype):
    for root, _folder, files in os.walk(logs_folder):
        print('\033[95m' + f'going through: {root}: {_folder} ' + '\033[0m')
        
        file_count = 0
        checkpoint = 0
        
        for file in files:
            if file_count == 0:
                print('machine progress%: ', end=" ")
            if file.endswith('.zip') and file != '.gitignore':
                zip_path = os.path.join(root, file)
                
                # Extract filename without extension
                filename = os.path.splitext(file)[0]
                
                # Extract Library and Machine from filename (e.g., YT-GFK-PAK1)
                library = filename.split('-')[0] if '-' in filename else filename
                machine = filename
                
                with tempfile.TemporaryDirectory() as temp_dir:
                    try:
                        # Extract zip file
                        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                            zip_ref.extractall(temp_dir)
                        
                        # Walk through the Log folder in the zip
                        log_dir = os.path.join(temp_dir, 'Log')
                        if not os.path.exists(log_dir):
                            print('\033[93m' + f"Log directory not found in zip file: {zip_path}\n\tgoing thru inside of the zip file" + '\033[95m')
                            recursive_walk_for_zip(temp_dir, log_filetype)
                            continue
                        if library in lib_machines_count:  
                            lib_machines_count[library] += 1
                        else:
                            lib_machines_count[library] = 1
                        lib_machines_count['all library'] += 1        
                        for log_root, _, log_files in os.walk(log_dir):
                            for log_file in log_files:
                                # Match date in filename and skip command.log
                                date_match = re.search(date_pattern, log_file)
                                if not date_match or log_file.endswith('command.log'):
                                    continue
                                
                                date_str = date_match.group(1)
                                
                                global start_date
                                global default_start_date
                                global end_date
                                global default_end_date
                                
                                new_date_str=datetime.datetime.strptime(date_str, '%Y-%m-%d')
                                
                                #check if have a default start date
                                if not start_date:
                                    start_date = default_start_date
                                #if both is false then no need do skipping
                                if start_date and not start_date  == "none":
                                    if( new_date_str < datetime.datetime.strptime(start_date, '%Y-%m-%d')):
                                        continue
                                
                                #check if have a default end date
                                if not end_date:
                                    end_date = default_end_date
                                #same same
                                if end_date and not end_date == "none":
                                    if( new_date_str > datetime.datetime.strptime(end_date, '%Y-%m-%d')):
                                        continue
                                
                                filepath = os.path.join(log_root, log_file)
                                
                                # Process both _local.log and .log files
                                if log_file.endswith(('_local.log', '.log')):
                                    log_data = process_log_file(filepath, date_str, library, machine)
                                    all_local_error_logs.extend(log_data)
                                    
                    except zipfile.BadZipFile:
                        print(f'\033[91m' + f"Invalid zip file: {zip_path}" + '\033[95m')
                        invalid_zip.append(zip_path)
                        continue
            file_count += 1
            progress = file_count / len(files) * 100
            
            if math.floor(progress) >= checkpoint:
                checkpoint += 10
                print('.', end=" ")   
        print('\033[93m' + '\nnum of library processed:', end=' ')
        print(len(lib_machines_count) - 1)
        print('num of library left:', end=' ')
        print(len(next(os.walk(logs_folder))[1]) - len(lib_machines_count) + 1)             
        print('\033[0m')
    with open('invalid_zip.txt', 'w') as f:
        count = 0
        for zip in invalid_zip:
            f.write(zip + '\n')
            count += 1
        print('\033[92m' + f"invalid_zip.txt created, total invalid zips: {count}")
# Regular expression to match date in filename (e.g., 2025-02-10)
date_pattern = r'(\d{4}-\d{2}-\d{2})'
# Initialize list to store all error data
all_local_error_logs = []
invalid_zip = []
lib_machines_count = {
        "all library": 0
    }

def extract_errors_to_single_excel(logs_folder='logs', output_excel='xlsx/error_logs.xlsx', log_filetype = 'local'):
    """Extract specified error types from local logs in all zip files across subfolders and save to a single Excel file."""
    
    recursive_walk_for_zip(logs_folder, log_filetype)
    
    # Create DataFrame
    local_error_df = pd.DataFrame(all_local_error_logs)
    
    # Reorder columns
    if not local_error_df.empty:
        error_cols = ['Library', 'Machine', 'Date', 'Time', 'Error Type']
        local_error_df = local_error_df[error_cols]
    
    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_excel), exist_ok=True)
    
    # Save to Excel
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        if not local_error_df.empty:
            local_error_df.to_excel(writer, sheet_name='Error_Logs', index=False)
            
            
    print(f"Excel file created successfully: {output_excel}" + '\033[0m')
    print('\033[94m' + f"number of machines per lib: {lib_machines_count}" + '\033[0m')
    

if __name__ == "__main__":
    #1
    #using current directory folder
    #RUN in local logs folder
    # folderpath = 'logs'   #UNCOMMENT to set this as default
    
    
    #2
    #using onedrive folder from C:\Users\isaacleong\Downloads\log_analyzer to C:\Users\isaacleong\WAFER SYSTEMS\Tin Lai - Log
    #RUN in relative path to one drive, may need to change depending where you downloaded this directory
    #------------------------------default settings-----------------------
    # folderpath = '../../WAFER SYSTEMS/Tin Lai - Log/30.6.2025/NPPL'    
    # today = datetime.datetime.now()
    # output_excel_location = today.strftime('xlsx/%d-%m-%Y_error_logs.xlsx')
    
    terminal_response = False
    while True: 
        print('\033[95m' + '-----------------------------------log analyzer-----------------------------------' + '\033[0m')
        if terminal_response:
            print(terminal_response)
            terminal_response = False
        else:
            print('')
        reply = input('\033[94m' + 'The current logs folder is: ' + '\033[4m' + folderpath + '\n' + '\033[0m' + '\033[94m' +
                      'The result will be stored to: ' + '\033[4m' + output_excel_location + '\n' + '\033[0m' + '\033[96m' +
                      'Commands:\n' +
                      '\tTo change get log location, please type \'logs\'\n' +
                      '\tTo change export excel file settings, please type \'export\'\n' +
                      '\tTo add or remove error types, please type \'errors\'\n' +
                      '\tTo view all info, please type \'info\'\n' +
                      '\tTo run the program, please type \'run\' \n' + '\033[0m'
                      )
        if reply == 'run':
            break
        elif reply == 'logs':
            print('\033[95m' + '-----------------------------------logs settings-----------------------------------\n' + '\033[0m')
            
            if start_date == False:
                current_start_date = default_start_date
            else:
                current_start_date = start_date
            
            if end_date == False:
                current_end_date = default_end_date
            else:
                current_end_date = end_date
            
            reply = input('\033[94m' + 'The current start date is: ' + str(current_start_date) + '\n' +
                          'The current end date is: ' + str(current_end_date) + '\n' +
                          '\033[96m' +'Commands:\n' +
                          '\tTo change start date, please type \'start yyyy-mm-dd\' \n' +
                          '\tTo change end date, please type \'end yyyy-mm-dd\' \n' +
                          '\tTo change path to using local directory logs, please type \'logs\' \n' +
                          '\tTo use another relative path / absolute path, please type the path directly\n' + 
                          '\tTo go back, please type \'back\'\n' + 
                          '\tTo run the program, please type \'run\' \n' + '\033[0m')
            if reply == 'back':
                continue
            if reply == 'run':
                break
            if reply.startswith('start'):
                words = re.split(r"\s", reply)
                if len(words) == 1:
                    terminal_response = '\033[91m' + 'Failure: Cannot change start date \n\tEmpty response: please enter a date' + '\033[0m'
                    continue
                if words[1] == 'none':
                    start_date = words[1]
                    terminal_response = '\033[92m' + 'start date changed successfully' + '\033[0m'
                    continue
                if re.match(date_pattern, words[1]):
                    if end_date and not end_date == "none":
                        if( datetime.datetime.strptime(words[1], '%Y-%m-%d') > datetime.datetime.strptime(end_date, '%Y-%m-%d')):
                            terminal_response = '\033[91m' + 'Failure: Cannot change end date \n\tStart date cannot be later than end date: please enter a valid date' + '\033[0m'
                            continue
                    start_date = words[1]
                    terminal_response = '\033[92m' + 'start date changed successfully' + '\033[0m'
                    continue
                else:
                    terminal_response = '\033[91m' + 'Failure: Cannot change start date \n\tIncompatible date: please enter a valid date' + '\033[0m'
                    continue
                
            if reply.startswith('end'):
                words = re.split(r"\s", reply)
                if len(words) == 1:
                    terminal_response = '\033[91m' + 'Failure: Cannot change end date \n\tEmpty response: please enter a date' + '\033[0m'
                    continue
                if words[1] == 'none':
                    end_date = words[1]
                    terminal_response = '\033[92m' + 'end date changed successfully' + '\033[0m'
                    continue
                if re.match(date_pattern, words[1]):
                    # #if both is false then no need do skipping
                    #             if start_date and not start_date  == "none":
                    #                 if( new_date_str < datetime.datetime.strptime(start_date, '%Y-%m-%d')):
                    #                     continue
                                
                    #             #check if have a default end date
                    #             if not end_date:
                    #                 end_date = default_end_date
                    #             #same same
                    #             if end_date and not end_date == "none":
                    #                 if( new_date_str > datetime.datetime.strptime(end_date, '%Y-%m-%d')):
                    #                     continue
                    if start_date and not start_date  == "none":    #checking if end earlier than start
                        if( datetime.datetime.strptime(words[1], '%Y-%m-%d') > datetime.datetime.strptime(start_date, '%Y-%m-%d')):
                            terminal_response = '\033[91m' + 'Failure: Cannot change end date \n\tEnd date cannot be earlier than start date: please enter a valid date' + '\033[0m'
                            continue
                    end_date = words[1]
                    terminal_response = '\033[92m' + 'end date changed successfully' + '\033[0m'
                    continue
                else:
                    terminal_response = '\033[91m' + 'Failure: Cannot change end date \n\tIncompatible date: please enter a valid date' + '\033[0m'
                    continue
                    
            folderpath = reply
            terminal_response = '\033[92m' + 'logs folder path changed successfully' + '\033[0m'
        elif reply == 'export':
            reply = input('\033[95m' + '-----------------------------------export settings-----------------------------------\n' + '\033[0m' + '\033[96m' +
                          'Default: ' + '\033[4m' + 'xlsx/<date>_error_logs.xlsx\n' + '\033[0m' + '\033[94m' +
                          'Current: ' + '\033[4m' + str(output_excel_location) + '\033[0m' + '\033[94m' +
                          'Commands:\n' +
                          '\tTo change path and filename, please type the full/relative path with correct filename\n' +
                          '\tTo go back, please type \'back\'\n' + 
                          '\tTo run the program, please type \'run\' \n' + '\033[0m')
            if reply == 'back':
                continue
            if reply == 'run':
                break
            if not re.search('.xlsx$', reply):
                terminal_response = '\033[91m' + 'Failure: Cannot change export filepath / filename: \n\tIncompatible filename: please end with a .xlsx' + '\033[0m'
                continue
            output_excel_location = re.sub('<date>', today.strftime('%d-%m-%Y'), reply)
            terminal_response = '\033[92m' + 'export path changed successfully' + '\033[0m'
            
        elif reply == 'errors':
            
            print('\033[95m' + '-----------------------------------errors settings-----------------------------------' + '\033[0m' + '\033[96m')
            data = get_error_json()
            for error in data:
                print('\033[94m' + error + ': ' + data[error] + '\033[96m')
            reply = input(
                          'Commands:\n' +
                          '\tTo reset to default, please type \'default\'\n' + 
                          '\tTo add error type, please type \'add <key> <error message>\', where error message do not contain any space\n' +
                          '\tTo remove error type, please type \'remove <key>\'\n' +
                          '\tTo go back, please type \'back\'\n' + 
                          '\tTo run the program, please type \'run\' \n' + '\033[0m')
            if reply == 'back':
                continue
            if reply == 'run':
                break
            if reply == 'default':
                os.remove("error_types.json")
                terminal_response = '\033[92m' + 'error types are set to default' + '\033[0m'
                continue
            if reply.startswith('add'):
                words = re.split(r"\s", reply)
                key = words[1]
                content = words[2]
                add_error_json(key, content)
                terminal_response = '\033[92m' + 'error type successfully added' + '\033[0m'
                continue
            if reply.startswith('remove'):
                words = re.split(r"\s", reply)
                key = words[1]
                remove_error_json(key)
                terminal_response = '\033[92m' + 'error type successfully removed' + '\033[0m'
                
        elif reply == 'info':
            if start_date == False:
                current_start_date = default_start_date
            else:
                current_start_date = start_date
            
            if end_date == False:
                current_end_date = default_end_date
            else:
                current_end_date = end_date
            print('\033[95m' + '-----------------------------------info-----------------------------------' + '\033[0m' + '\033[96m')
            print('\033[94m' + 'The current start date is: ' + str(current_start_date) + ' (logs)\n' +
                          'The current end date is: ' + str(current_end_date) + ' (logs)\n\n' +
                          'Default export path: ' + '\033[4m' + 'xlsx/<date>_error_logs.xlsx (logs)\n' + '\033[0m' + '\033[94m' +
                          'Current export path: ' + '\033[4m' + str(output_excel_location) + ' (logs)\033[0m' + '\033[94m\n\n' + 
                          'Default: ' + '\033[4m' + 'xlsx/<date>_error_logs.xlsx(export)\n' + '\033[0m' + '\033[94m' +
                          'Current: ' + '\033[4m' + str(output_excel_location) + '(export)\033[0m' + '\033[94m\n')
            data = get_error_json()
            for error in data:
                print('\033[94m' + error + ': ' + data[error] + '(errors)\033[96m')
            reply = input('Commands:\n' +
                          '\tTo go back, please type \'back\'\n' + 
                          '\tTo run the program, please type \'run\' \n' + '\033[0m')
            if reply == 'back':
                continue
            if reply == 'run':
                break
        else: 
            terminal_response = '\033[93m' + 'Warning: this command do not exist' + '\033[0m'
            continue
            
        
    extract_errors_to_single_excel(logs_folder=folderpath, output_excel = output_excel_location)