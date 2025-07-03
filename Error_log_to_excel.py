import zipfile
import os
import re
import pandas as pd
import tempfile
import datetime
from pathlib import Path
import json
import math

def parse_local_log_line(line):
    """Parse a local log line to extract timestamp and message."""
    pattern = r'(\d{2}:\d{2}:\d{2})\s+\d+\s+\[[^\]]+\](.*)'
    match = re.match(pattern, line.strip())
    if match:
        return {
            'Time': match.group(1),
            'Message': match.group(2).strip()
        }
    return None

def getJSONFILE(path):
    datafile = path
    try:
        if datafile.is_file():    
            with open(datafile, encoding='utf-8') as json_file:
                data = json.load(json_file)
        else:
            data = {
                "A1": "检测服务器状态：False",
                "A2": "程序启动时，网络异常，将进入离线模式",
                'B1': "获取工作站状态失败，服务不可访问",
                "B2": "登录页，已到最大重试次数，进入离线模式"
            }
            with open(datafile, 'w') as json_file:
                json.dump(data, json_file, indent=4)
    except json.JSONDecodeError as e:
        print(f"JSONDecodeError: {e}. Initializing with default data.")
        data = {
            "A1": "检测服务器状态：False",
            "A2": "程序启动时，网络异常，将进入离线模式",
            'B1': "获取工作站状态失败，服务不可访问",
            "B2": "登录页，已到最大重试次数，进入离线模式"
        }
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
        
def get_error_type(message):
    # error_types = {
    #     'A1': r'检测服务器状态：False',
    #     'A2': r'程序启动时，网络异常，将进入离线模式',
    #     'B1': r'获取工作站状态失败，服务不可访问',
    #     'B2': r'登录页，已到最大重试次数，进入离线模式'
    # }
    error_types = get_error_json()
    """Determine the error type based on the message content."""
    for error in error_types:
        if re.search(error_types[error], message):
            return error
    return None

def process_local_log_file(filepath, date_str, library, machine):
    """Read a local log file and extract error-related data."""
    data = []
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            parsed = parse_local_log_line(line)
            if parsed:
                # error_type = get_error_type(parsed['Message'])
                error_type = get_error_type(line)
                if error_type:
                    parsed['Error Type'] = error_type
                    parsed['Date'] = date_str
                    parsed['Library'] = library
                    parsed['Machine'] = machine
                    data.append(parsed)
    return data


def recursive_walk_for_zip(logs_folder):
    folder_count = 0
    for root, _folder, files in os.walk(logs_folder):
        print('\033[95m'+f'\ngoing through: {root}: {_folder} '+'\033[0m')
        
        file_count = 0
        checkpoint = 0
        
        for file in files:
            if file_count == 0:
                print('num of library processed:', end=' ')
                print(len(lib_machines_count) - 1)
                print('num of library left:', end=' ')
                print(len(next(os.walk(logs_folder))[1]) - len(lib_machines_count) + 1)
                print('machine progress%: ', end=" ")
            if file.endswith('.zip') and file != '.gitignore':
                zip_path = os.path.join(root, file)
                
                # Extract filename without extension
                filename = os.path.splitext(file)[0]
                
                # Extract Library and Machine from filename (e.g., YT-GFK-PAK1)
                library = filename.split('-')[0] if '-' in filename else filename
                machine = filename
                
                # if library in lib_machines_count:  
                #     lib_machines_count[library] += 1
                # else:
                #     lib_machines_count[library] = 1
                # lib_machines_count['all library'] += 1
                # Create a temporary directory to extract files
                with tempfile.TemporaryDirectory() as temp_dir:
                    try:
                        # Extract zip file
                        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                            zip_ref.extractall(temp_dir)
                        
                        # Walk through the Log folder in the zip
                        log_dir = os.path.join(temp_dir, 'Log')
                        if not os.path.exists(log_dir):
                            print(f'\033[93m' + "Log directory not found in zip file: {zip_path}\n\tgoing thru inside of the zip file" + '\033[95m')
                            recursive_walk_for_zip(temp_dir)
                            continue
                        if library in lib_machines_count:  
                            lib_machines_count[library] += 1
                        else:
                            lib_machines_count[library] = 1
                        lib_machines_count['all library'] += 1        
                        for log_root, _, log_files in os.walk(log_dir):
                            for log_file in log_files:
                                # Match date in filename and process only _local.log files
                                date_match = re.search(date_pattern, log_file)
                                if not date_match or not log_file.endswith('_local.log'):
                                    continue
                                date_str = date_match.group(1)
                                
                                # Process local log file
                                filepath = os.path.join(log_root, log_file)
                                log_data = process_local_log_file(filepath, date_str, library, machine)
                                all_error_logs.extend(log_data)
                    except zipfile.BadZipFile:
                        print(f'\033[91m' + "Invalid zip file: {zip_path}" + '\033[95m')
                        invalid_zip.append(zip_path)
                        continue
            file_count += 1
            progress = file_count / len(files) * 100
            
            if (math.floor(progress) >= checkpoint):
                checkpoint = checkpoint + 10
                print('.', end=" ")                
# Regular expression to match date in filename (e.g., 2025-02-10)
date_pattern = r'(\d{4}-\d{2}-\d{2})'
# Initialize list to store all error data
all_error_logs = []
invalid_zip = []
lib_machines_count = {
        "all library": 0
    }

def extract_errors_to_single_excel(logs_folder='logs', output_excel='xlsx/error_logs.xlsx'):
    """Extract specified error types from local logs in all zip files across subfolders and save to a single Excel file."""
    
    
    # Walk through logs folder and all subfolders
    for root, _folder, files in os.walk(logs_folder):
        print('\033[95m'+f'\ngoing through: {root}: {_folder} '+'\033[0m')

        file_count = 0
        checkpoint = 0
        for file in files:
            if file_count == 0:
                print('num of library processed:', end=' ')
                print(len(lib_machines_count) - 1)
                print('num of library left:', end=' ')
                print(len(next(os.walk(logs_folder))[1]) - len(lib_machines_count) + 1)
                print('machine progress%: ', end=" ")
            if file.endswith('.zip') and file != '.gitignore':
                zip_path = os.path.join(root, file)
                
                # Extract filename without extension
                filename = os.path.splitext(file)[0]
                
                # Extract Library and Machine from filename (e.g., YT-GFK-PAK1)
                library = filename.split('-')[0] if '-' in filename else filename
                machine = filename
                
                # Create a temporary directory to extract files
                with tempfile.TemporaryDirectory() as temp_dir:
                    try:
                        # Extract zip file
                        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                            zip_ref.extractall(temp_dir)
                        
                        # Walk through the Log folder in the zip
                        log_dir = os.path.join(temp_dir, 'Log')
                        if not os.path.exists(log_dir):
                            print('\033[93m' + f"Log directory not found in zip file: {zip_path}\n\tgoing thru inside of the zip file" + '\033[95m')
                            recursive_walk_for_zip(temp_dir)
                            continue
                        if library in lib_machines_count:  
                            lib_machines_count[library] += 1
                        else:
                            lib_machines_count[library] = 1
                        lib_machines_count['all library'] += 1        
                        for log_root, _, log_files in os.walk(log_dir):
                            for log_file in log_files:
                                # Match date in filename and process only _local.log files
                                date_match = re.search(date_pattern, log_file)
                                if not date_match or not log_file.endswith('_local.log'):
                                    continue
                                date_str = date_match.group(1)
                                
                                # Process local log file
                                filepath = os.path.join(log_root, log_file)
                                log_data = process_local_log_file(filepath, date_str, library, machine)
                                all_error_logs.extend(log_data)
                    except zipfile.BadZipFile:
                        print(f'\033[91m' + "Invalid zip file: {zip_path}" + '\033[95m')
                        invalid_zip.append(zip_path)
                        continue
            file_count += 1
            progress = file_count / len(files) * 100
            
            if (math.floor(progress) >= checkpoint):
                checkpoint = checkpoint + 10
                print('.', end=" ")
                
    
    with open('invalid_zip.txt', 'w') as f:
        count = 0
        for zip in invalid_zip:
            f.write(zip + '\n')
            count += 1
        print('\033[92m' + f"invalid_zip.txt created, total invalid zips: {count}")
    # Create DataFrame
    error_df = pd.DataFrame(all_error_logs)
    
    # Reorder columns
    if not error_df.empty:
        error_cols = ['Library', 'Machine', 'Date', 'Time', 'Error Type']
        error_df = error_df[error_cols]
    
    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_excel), exist_ok=True)
    
    # Save to Excel
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        if not error_df.empty:
            error_df.to_excel(writer, sheet_name='Error_Logs', index=False)
            
            
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
    folderpath = '../../WAFER SYSTEMS/Tin Lai - Log/30.6.2025'     #UNCOMMENT to set this as default
    today = datetime.datetime.now()
    output_excel_location = today.strftime('xlsx/%d-%m-%Y_error_logs.xlsx')
    
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
                      '\tTo run the program, please type \'run\' \n' + '\033[0m'
                      )
        if reply == 'run':
            break
        elif reply == 'logs':
            reply = input('\033[95m' + '-----------------------------------logs settings-----------------------------------\n' + '\033[0m' + '\033[96m' +
                          'Commands:\n' +
                          '\tTo change path to using local directory logs, please type \'logs\' \n' +
                          '\tTo use another relative path / absolute path, please type the path directly\n' + 
                          '\tTo go back, please type \'back\'\n' + 
                          '\tTo run the program, please type \'run\' \n' + '\033[0m')
            if reply == 'back':
                continue
            if reply == 'run':
                break
            folderpath = reply
            terminal_response = '\033[92m' + 'logs folder path changed successfully' + '\033[0m'
        elif reply == 'export':
            reply = input('\033[95m' + '-----------------------------------export settings-----------------------------------\n' + '\033[0m' + '\033[96m' +
                          'Default: ' + '\033[4m' + 'xlsx/<date>_error_logs.xlsx\n' + '\033[0m' + '\033[94m' +
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
            data = get_error_json()
            print('\033[95m' + '-----------------------------------errors settings-----------------------------------' + '\033[0m' + '\033[96m')
            for error in data:
                print('\033[94m' + error + ': ' + data[error] + '\033[96m')
            reply = input(
                          'Commands:\n' +
                          '\tTo add error type, please type \'add <key> <error message>\', where error message do not contain any space\n' +
                          '\tTo remove error type, please type \'remove <key>\'\n' +
                          '\tTo go back, please type \'back\'\n' + 
                          '\tTo run the program, please type \'run\' \n' + '\033[0m')
            if reply == 'back':
                continue
            if reply == 'run':
                break
            if reply.startswith('add'):
                words = re.split(r"\s", reply)
                key = words[1]
                content = words[2]
                add_error_json(key, content)
                continue
            if reply.startswith('remove'):
                words = re.split(r"\s", reply)
                key = words[1]
                remove_error_json(key)
            
        else: 
            terminal_response = '\033[93m' + 'Warning: this command do not exist' + '\033[0m'
            continue
            
        
    extract_errors_to_single_excel(logs_folder=folderpath, output_excel = output_excel_location)