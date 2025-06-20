import zipfile
import os
import re
import pandas as pd
import tempfile

def parse_local_log_line(line):
    """Parse a local log line to extract timestamp and message."""
    # Example: 09:24:12 645 [ServiceUnavailableHelper.CheckServerConnect]服务器状态：False
    pattern = r'(\d{2}:\d{2}:\d{2})\s+\d+\s+\[[^\]]+\](.*)'
    match = re.match(pattern, line.strip())
    if match:
        return {
            'Time': match.group(1),
            'Message': match.group(2).strip()
        }
    return None

def get_error_type(message):
    """Determine the error type based on the message content."""
    if message == '检测服务器状态：False':
        return 'A1'
    elif message == '程序启动时，网络异常，将进入离线模式':
        return 'A2'
    elif message == '获取工作站状态失败，服务不可访问':
        return 'B1'
    elif message == '登录页，已到最大重试次数，进入离线模式':
        return 'B2'
    return None

def process_local_log_file(filepath, date_str, library, machine):
    """Read a local log file and extract error-related data."""
    data = []
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            parsed = parse_local_log_line(line)
            if parsed:
                error_type = get_error_type(parsed['Message'])
                if error_type:
                    parsed['Error Type'] = error_type
                    parsed['Date'] = date_str
                    parsed['Library'] = library
                    parsed['Machine'] = machine
                    data.append(parsed)
    return data

def extract_errors_to_single_excel(logs_folder='logs', output_excel='xlsx/error_logs.xlsx'):
    """Extract specified error types from local logs in all zip files across subfolders and save to a single Excel file."""
    # Initialize list to store all error data
    all_error_logs = []
    
    # Regular expression to match date in filename (e.g., 2025-02-10)
    date_pattern = r'(\d{4}-\d{2}-\d{2})'
    
    # Walk through logs folder and all subfolders
    for root, _folder, files in os.walk(logs_folder):
        print(f'going through: {_folder} ... ')
        for file in files:
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
                            print(f"Log directory not found in zip file: {zip_path}")
                            continue
                        
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
                        print(f"Invalid zip file: {zip_path}")
                        continue
    
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
    
    print(f"Excel file created successfully: {output_excel}")

if __name__ == "__main__":
    # Example usage
    extract_errors_to_single_excel(logs_folder='logs', output_excel='xlsx/error_logs.xlsx')