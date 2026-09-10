import csv
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import sys
import os
import time
import concurrent.futures
import threading

# Global flag to signal cancellation/exit
exit_requested = threading.Event()
prompt_active = threading.Event()

def input_listener():
    """Background thread that listens for user input to exit ('x') once activated."""
    while not exit_requested.is_set():
        try:
            line = sys.stdin.readline()
            if line and line.strip().lower() == 'x':
                print("\n[INFO] 'x' entered by user. Terminating program...")
                os._exit(0)
        except Exception:
            break

def check_name_column(col_info, sheet):
    col_name, col_idx = col_info
    if col_name and str(col_name).endswith("Name"):
        unique_values = set()
        for row_idx in range(2, sheet.max_row + 1):
            cell_value = sheet.cell(row=row_idx, column=col_idx).value
            if cell_value is not None and str(cell_value).strip().lower() != "null":
                unique_values.add(str(cell_value).strip())
        
        if len(unique_values) == 1:
            return col_idx, col_name, unique_values.pop()
    return None

def check_null_column(col_idx, sheet, columns_to_delete):
    if col_idx in columns_to_delete:
        return None
    is_null_column = True
    for row_idx in range(2, sheet.max_row + 1):
        cell_value = sheet.cell(row=row_idx, column=col_idx).value
        if cell_value is not None and str(cell_value).strip().lower() != "null" and str(cell_value).strip() != "":
            is_null_column = False
            break
    if is_null_column:
        return col_idx
    return None

def remove_null_columns(filename):
    start_time = time.time()  # Record start time
    
    # Check the file extension
    ext = os.path.splitext(filename)[1].lower()
    
    try:
        if ext == '.csv':
            wb = openpyxl.Workbook()
            sheet = wb.active
            with open(filename, 'r', encoding='utf-8', errors='ignore') as f:
                reader = csv.reader(f)
                for r_idx, row in enumerate(reader, 1):
                    for c_idx, val in enumerate(row, 1):
                        sheet.cell(row=r_idx, column=c_idx, value=val)
        elif ext == '.xlsx':
            wb = openpyxl.load_workbook(filename)
            sheet = wb.active
        else:
            print("Error: Unsupported file type. Only .csv and .xlsx files are allowed.")
            return
            
    except FileNotFoundError:
        print(f"Error: The specified file '{filename}' was not found.")
        return
    except Exception as e:
        print(f"Error: Failed to load file:\n{str(e)}")
        return

    columns_to_delete = set()

    # --- 1. REARRANGE COLUMNS: rawEvent (1st) and eventTime (2nd) ---
    initial_positions = {cell.value: cell.column for cell in sheet[1] if cell.value is not None}
    
    if 'rawEvent' in initial_positions:
        raw_event_idx = initial_positions['rawEvent']
        sheet.insert_cols(1)
        for row_idx in range(1, sheet.max_row + 1):
            sheet.cell(row=row_idx, column=1, value=sheet.cell(row=row_idx, column=raw_event_idx + 1).value)
        sheet.delete_cols(raw_event_idx + 1)
        
    mid_positions = {sheet.cell(row=1, column=col).value: col for col in range(1, sheet.max_column + 1)}
    
    if 'eventTime' in mid_positions:
        event_time_idx = mid_positions['eventTime']
        sheet.insert_cols(2)
        for row_idx in range(1, sheet.max_row + 1):
            sheet.cell(row=row_idx, column=2, value=sheet.cell(row=row_idx, column=event_time_idx + 1).value)
        sheet.delete_cols(event_time_idx + 1)

    current_positions = {sheet.cell(row=1, column=col).value: col for col in range(1, sheet.max_column + 1)}

    delete_columns_names = ['rawEventHash', 'aisaacReceivedTime', 'customerURI', 'cenNifiReceiptTime', 'logFilterKafkaInTime', 'logFilterInTime']
    for name in delete_columns_names:
        if name in current_positions:
            columns_to_delete.add(current_positions[name])

    # --- 3. Multithreaded parsing of "Name" metadata columns ---
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [executor.submit(check_name_column, item, sheet) for item in current_positions.items()]
        for future in concurrent.futures.as_completed(futures):
            res = future.result()
            if res:
                col_idx, col_name, unique_value = res
                target_header_name = col_name.replace("Name", "")
                if target_header_name in current_positions:
                    target_col_idx = current_positions[target_header_name]
                    sheet.cell(row=1, column=target_col_idx, value=unique_value)
                    columns_to_delete.add(col_idx)

    # --- 4. Multithreaded check for empty or purely string "null" columns ---
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [executor.submit(check_null_column, col_idx, sheet, columns_to_delete) for col_idx in range(1, sheet.max_column + 1)]
        for future in concurrent.futures.as_completed(futures):
            res = future.result()
            if res is not None:
                columns_to_delete.add(res)

    # --- 5. Delete flagged columns ---
    for col_index in sorted(list(columns_to_delete), reverse=True):
        sheet.delete_cols(col_index)

    # Style formatting for the header row
    for cell in sheet[1]:
        cell.border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'),
                                           right=openpyxl.styles.Side(style='thin'),
                                           top=openpyxl.styles.Side(style='thin'),
                                           bottom=openpyxl.styles.Side(style='thin'))
        cell.font = Font(name='Calibri', size=13, bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Style formatting for data rows
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'),
                                               right=openpyxl.styles.Side(style='thin'),
                                               top=openpyxl.styles.Side(style='thin'),
                                               bottom=openpyxl.styles.Side(style='thin'))
            cell.font = Font(name='Calibri', size=12, color='000000')

    # Adjust column width dynamically
    for column_cells in sheet.columns:
        max_length = 0
        column = openpyxl.utils.get_column_letter(column_cells[0].column)
        col_header = column_cells[0].value
        
        for cell in column_cells:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        adjusted_width = (max_length + 2) * 1.2
        
        if col_header == 'rawEvent':
            sheet.column_dimensions[column].width = 50
        else:
            sheet.column_dimensions[column].width = max(min(adjusted_width, 40), 10)

    # Save workbook
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    base_path, _ = os.path.splitext(filename)
    output_filename = f"{base_path}_{timestamp}.xlsx"
    
    try:
        wb.save(output_filename)
        elapsed_time = time.time() - start_time
        print("\n[SUCCESS] ----------------------------------------")
        print(f"Modified workbook saved as: {output_filename}")
        print(f"Elapsed Time: {elapsed_time:.2f} seconds")
        print("--------------------------------------------------")
    except PermissionError:
        print("[ERROR] Permission denied. Please make sure write access is available.")

if __name__ == "__main__":
    print("=== Log Normalizer (Multithreaded with 60s Timeout Watcher) ===")
    print("Developed by Gourab Sarkar\n")

    if len(sys.argv) > 1:
        filename = sys.argv[1]
    else:
        filename = input("Enter the path to your CSV or Excel file: ").strip().strip('"').strip("'")

    if not filename:
        print("Error: No file path provided.")
        sys.exit(1)

    ext = os.path.splitext(filename)[1].lower()
    if ext not in ['.xlsx', '.csv']:
        print(f"Error: Invalid file type '{ext}'. Only .csv and .xlsx input files are allowed!")
        sys.exit(1)

    # Start the background listener thread for keyboard input
    listener_thread = threading.Thread(target=input_listener, daemon=True)
    listener_thread.start()

    print("[INFO] Processing file... Please wait.")

    # Run processing in a worker thread so we can monitor time
    worker_thread = threading.Thread(target=remove_null_columns, args=(filename,))
    worker_thread.start()

    start_wait = time.time()
    warned_60s = False

    while worker_thread.is_alive():
        worker_thread.join(timeout=1.0) # Check every 1 second
        elapsed = time.time() - start_wait
        if elapsed > 60 and not warned_60s:
            print("\n[NOTICE] Program has been running for over 60 seconds.")
            print("If you want to exit the program, enter 'x' and press Enter: ", end="", flush=True)
            warned_60s = True

    exit_requested.set()