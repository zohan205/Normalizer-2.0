import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import sys
import os
import time
import threading

exit_requested = threading.Event()

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

def process_log_file(filename):
    total_start = time.time()
    ext = os.path.splitext(filename)[1].lower()
    
    # --- 1. Load File using Pandas ---
    t0 = time.time()
    try:
        if ext == '.csv':
            df = pd.read_csv(filename, low_memory=False)
        elif ext == '.xlsx':
            df = pd.read_excel(filename, engine='openpyxl')
        else:
            print("Error: Unsupported file type.")
            return
    except Exception as e:
        print(f"Error loading file: {e}")
        return
    print(f"[DEBUG] 1. Pandas File Load Time: {time.time() - t0:.2f} seconds")

    # --- 2. Rearrange Columns ('rawEvent' 1st, 'eventTime' 2nd) ---
    t0 = time.time()
    cols = list(df.columns)
    
    # Pull rawEvent and eventTime to the front if they exist
    for col in ['eventTime', 'rawEvent']:
        if col in cols:
            cols.remove(col)
            cols.insert(0, col)
    
    # Ensure rawEvent is strictly 1st and eventTime is strictly 2nd if both exist
    if 'rawEvent' in df.columns and 'eventTime' in df.columns:
        cols = [c for c in cols if c not in ['rawEvent', 'eventTime']]
        cols = ['rawEvent', 'eventTime'] + cols

    df = df[cols]
    print(f"[DEBUG] 2. Column Rearranging Time: {time.time() - t0:.2f} seconds")

    # --- 3. Drop Pre-defined Unwanted Columns ---
    drop_columns_names = ['rawEventHash', 'aisaacReceivedTime', 'customerURI', 'cenNifiReceiptTime', 'logFilterKafkaInTime', 'logFilterInTime']
    df.drop(columns=[c for c in drop_columns_names if c in df.columns], inplace=True)

    # --- 4. Metadata "Name" Column Check & Rename Header ---
    t0 = time.time()
    cols_to_drop = []
    for col in df.columns:
        if str(col).endswith("Name"):
            target_header = col[:-4] # Remove "Name"
            if target_header in df.columns:
                # Check unique non-null/non-null-string values
                series = df[col].astype(str).str.strip()
                valid_mask = (df[col].notna()) & (series.str.lower() != 'null') & (series != '')
                unique_vals = series[valid_mask].unique()
                
                if len(unique_vals) == 1:
                    single_val = unique_vals[0]
                    # Update target column header/name (in pandas we can rename the column)
                    # Wait, the original script replaced the *header cell value* of the target column with the unique value.
                    # Let's rename the target column to the unique value!
                    df.rename(columns={target_header: single_val}, inplace=True)
                    cols_to_drop.append(col)
                    
    if cols_to_drop:
        df.drop(columns=cols_to_drop, inplace=True)
    print(f"[DEBUG] 3. Metadata Name Check Time: {time.time() - t0:.2f} seconds")

    # --- 5. Null Column Scan (Drop empty or purely 'null' columns) ---
    t0 = time.time()
    null_cols_to_drop = []
    for col in df.columns:
        series = df[col].astype(str).str.strip().str.lower()
        # Check if there are any non-empty, non-null values
        has_valid_data = (df[col].notna()) & (series != '') & (series != 'null') & (series != 'nan')
        if not has_valid_data.any():
            null_cols_to_drop.append(col)
            
    if null_cols_to_drop:
        df.drop(columns=null_cols_to_drop, inplace=True)
    print(f"[DEBUG] 4. Null Column Scan Time: {time.time() - t0:.2f} seconds")

    # --- 6. Save to Excel First ---
    t0 = time.time()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    base_path, _ = os.path.splitext(filename)
    output_filename = f"{base_path}_{timestamp}.xlsx"
    df.to_excel(output_filename, index=False, engine='openpyxl')
    print(f"[DEBUG] 5. Pandas Excel Write Time: {time.time() - t0:.2f} seconds")

    # --- 7. Apply Fast OpenPyXL Styling on the Smaller Cleaned Workbook ---
    t0 = time.time()
    wb = openpyxl.load_workbook(output_filename)
    sheet = wb.active

    thin_border = openpyxl.styles.Border(
        left=openpyxl.styles.Side(style='thin'), right=openpyxl.styles.Side(style='thin'),
        top=openpyxl.styles.Side(style='thin'), bottom=openpyxl.styles.Side(style='thin')
    )
    header_font = Font(name='Calibri', size=13, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
    header_align = Alignment(horizontal='center', vertical='center')

    # Style Header
    for cell in sheet[1]:
        cell.border = thin_border
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    # Style Rows & Auto-adjust Widths
    row_font = Font(name='Calibri', size=12, color='000000')
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.border = thin_border
            cell.font = row_font

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

    wb.save(output_filename)
    print(f"[DEBUG] 6. Styling & Width Adjustment Time: {time.time() - t0:.2f} seconds")

    print(f"\n[SUCCESS] Total Elapsed Time: {time.time() - total_start:.2f} seconds")
    print(f"Saved as: {output_filename}")

if __name__ == "__main__":
    print("=== Log Normalizer (Pandas Accelerated Engine) ===")
    filename = sys.argv[1] if len(sys.argv) > 1 else input("Enter file path: ").strip().strip('"').strip("'")
    
    if not filename or not os.path.exists(filename):
        print(f"Error: File '{filename}' not found.")
        sys.exit(1)

    # Start listener thread for 'x' cancellation option
    threading.Thread(target=input_listener, daemon=True).start()

    # Run processing in a background worker thread
    worker = threading.Thread(target=process_log_file, args=(filename,))
    worker.start()

    start_wait = time.time()
    warned = False
    while worker.is_alive():
        worker.join(timeout=1.0)
        if time.time() - start_wait > 60 and not warned:
            print("\n[NOTICE] Program has been running for over 60 seconds.")
            print("If you want to exit the program, enter 'x' and press Enter: ", end="", flush=True)
            warned = True
            
    exit_requested.set()