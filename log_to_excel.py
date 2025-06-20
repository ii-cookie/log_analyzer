import zipfile
import os
import re
import pandas as pd
from datetime import datetime
import tempfile

def parse_log_line(line, log_type):
    """Parse a single log line based on log file type and return a dictionary of extracted fields."""
    if log_type == "main":
        # Example: 10:20:28.389 [Info]启动MQTT重连定时器
        pattern = r'(\d{2}:\d{2}:\d{2}\.\d{3})\s*\[([^\]]+)\](.*)'
        match = re.match(pattern, line.strip())
        if match:
            return {
                'Timestamp': match.group(1),
                'LogLevel': match.group(2),
                'Message': match.group(3).strip()
            }
    elif log_type == "local":
        # Example: 09:24:12 645 [ServiceUnavailableHelper.CheckServerConnect]服务器状态：False
        pattern = r'(\d{2}:\d{2}:\d{2})\s+(\d+)\s+\[([^\]]+)\](.*)'
        match = re.match(pattern, line.strip())
        if match:
            return {
                'Timestamp': match.group(1),
                'Sequence': match.group(2),
                'Component': match.group(3),
                'Message': match.group(4).strip()
            }
    elif log_type == "command":
        # Example: 09:23:12.395  HttpHelper Send   https://lt.slsnp.hkpl.gov.hk/api/cfm-admin-service/portal-api/clientParameter
        pattern = r'(\d{2}:\d{2}:\d{2}\.\d{3})\s+([^\s]+)\s+([^\s]+)\s*(.*)'
        match = re.match(pattern, line.strip())
        if match:
            return {
                'Timestamp': match.group(1),
                'Component': match.group(2),
                'Action': match.group(3),
                'Details': match.group(4).strip()
            }
    return None

def process_log_file(filepath, log_type):
    """Read a log file and extract data into a list of dictionaries."""
    data = []
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            parsed = parse_log_line(line, log_type)
            if parsed:
                data.append(parsed)
    return data

def extract_logs_to_excel(zip_path, output_excel):
    """Extract logs from a zip file and save to an Excel file."""
    # Create a temporary directory to extract files
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract zip file
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Initialize dictionaries to store data for each log type
        main_logs = []
        local_logs = []
        command_logs = []
        
        # Walk through the Log folder
        log_dir = os.path.join(temp_dir, 'Log')
        if not os.path.exists(log_dir):
            print(f"Log directory not found in zip file: {log_dir}")
            return
        
        # Regular expression to match date in filename (e.g., 2025-02-10)
        date_pattern = r'(\d{4}-\d{2}-\d{2})'
        
        for root, _, files in os.walk(log_dir):
            for file in files:
                # Match date in filename
                date_match = re.search(date_pattern, file)
                if not date_match:
                    continue
                date_str = date_match.group(1)
                
                # Determine log type based on filename
                filepath = os.path.join(root, file)
                if file.endswith('.log') and not file.endswith('_local.log') and not file.endswith('_command.log'):
                    log_data = process_log_file(filepath, 'main')
                    for entry in log_data:
                        entry['Date'] = date_str
                    main_logs.extend(log_data)
                elif file.endswith('_local.log'):
                    log_data = process_log_file(filepath, 'local')
                    for entry in log_data:
                        entry['Date'] = date_str
                    local_logs.extend(log_data)
                elif file.endswith('_command.log'):
                    log_data = process_log_file(filepath, 'command')
                    for entry in log_data:
                        entry['Date'] = date_str
                    command_logs.extend(log_data)
        
        # Create DataFrames
        main_df = pd.DataFrame(main_logs)
        local_df = pd.DataFrame(local_logs)
        command_df = pd.DataFrame(command_logs)
        
        # Reorder columns to have Date and Timestamp first
        if not main_df.empty:
            main_cols = ['Date', 'Timestamp', 'LogLevel', 'Message']
            main_df = main_df[main_cols]
        if not local_df.empty:
            local_cols = ['Date', 'Timestamp', 'Sequence', 'Component', 'Message']
            local_df = local_df[local_cols]
        if not command_df.empty:
            command_cols = ['Date', 'Timestamp', 'Component', 'Action', 'Details']
            command_df = command_df[command_cols]
        
        # Save to Excel with separate sheets
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            if not main_df.empty:
                main_df.to_excel(writer, sheet_name='Main_Logs', index=False)
            if not local_df.empty:
                local_df.to_excel(writer, sheet_name='Local_Logs', index=False)
            if not command_df.empty:
                command_df.to_excel(writer, sheet_name='Command_Logs', index=False)
        
        print(f"Excel file created successfully: {output_excel}")

if __name__ == "__main__":
    # Example usage
    zip_file_path = "POR-01C-PDC7.zip"  # Replace with your zip file path
    output_excel_path = "log_output.xlsx"  # Desired output Excel file
    extract_logs_to_excel(zip_file_path, output_excel_path)