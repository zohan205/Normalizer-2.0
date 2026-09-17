import os
import sys
import time
import threading
from datetime import datetime

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# ============================================================
# APPLICATION STATE (SOC DARK THEME)
# ============================================================

APP_TITLE = "Formatix...a Log Normalizer"
APP_BACKGROUND = "#0F172A"       # Deep Obsidian / Dark Slate
APP_CARD_BG = "#1E293B"          # Dark Panel Card Color
APP_ACCENT = "#00F0FF"           # Cyber Cyan Highlight
APP_BTN_PRIMARY = "#0EA5E9"      # Tactical Blue Button
APP_BTN_HOVER = "#0284C7"        # Active State Cyan-Blue
APP_TEXT_MAIN = "#F8FAFC"        # Bright White Text
APP_TEXT_MUTED = "#94A3B8"       # Subtitle Gray Text
APP_SUCCESS = "#10B981"          # Emerald Neon Green
APP_ERROR = "#EF4444"            # Alert Red
APP_WARNING = "#F59E0B"          # Warning Amber

processing_active = False
processing_started_at = None
selected_filename = ""
last_output_filename = ""


# ============================================================
# GUI-SAFE HELPERS
# ============================================================

def write_log(message):
    pass

def _write_log_now(message):
    pass

def set_status(message, color=APP_TEXT_MUTED):
    """Update the status label safely from any thread."""
    root.after(0, lambda: status_label.config(text=message, fg=color))


def set_stage(message):
    """Update both the status label and activity log."""
    set_status(message, APP_TEXT_MUTED)
    write_log(message)


def clean_display_path(path):
    """Return a readable absolute path."""
    return os.path.abspath(os.path.expanduser(str(path)))


# ============================================================
# DATA CLEANING FUNCTIONS
# ============================================================

def clean_name_series(series):
    cleaned_series = series.astype("string").str.strip()
    lowercase_series = cleaned_series.str.lower()

    valid_mask = (
        series.notna()
        & cleaned_series.notna()
        & cleaned_series.ne("")
        & lowercase_series.ne("null")
        & lowercase_series.ne("nan")
        & lowercase_series.ne("none")
    )

    return cleaned_series, valid_mask


def clean_filename_column_name(column_name):
    column_name = str(column_name).strip()
    column_name = column_name.replace("\r", " ")
    column_name = column_name.replace("\n", " ")
    column_name = column_name.replace("\t", " ")
    return column_name


def get_safe_output_column_name(
    requested_name,
    original_columns,
    generated_columns,
    source_value_column
):
    requested_name = clean_filename_column_name(requested_name)

    if not requested_name:
        return None

    if requested_name in generated_columns:
        return requested_name

    if requested_name not in original_columns:
        return requested_name

    if requested_name == source_value_column:
        return requested_name

    counter = 2
    safe_name = f"{requested_name}_{counter}"
    all_used_columns = set(original_columns) | set(generated_columns)

    while safe_name in all_used_columns:
        counter += 1
        safe_name = f"{requested_name}_{counter}"

    return safe_name


def expand_metadata_columns_fast(df):
    original_index = df.index
    original_row_count = len(df)
    original_columns = list(df.columns)
    original_column_set = set(original_columns)

    metadata_pairs = []

    for name_column in original_columns:
        name_column_text = str(name_column)

        if not name_column_text.endswith("Name"):
            continue

        value_column = name_column_text[:-4]

        if value_column in df.columns:
            metadata_pairs.append((name_column, value_column))

    if not metadata_pairs:
        return df

    generated_columns = {}
    columns_to_drop = []

    for name_column, value_column in metadata_pairs:
        cleaned_names, valid_name_mask = clean_name_series(df[name_column])

        unique_names = (
            cleaned_names.loc[valid_name_mask]
            .drop_duplicates()
            .tolist()
        )

        for unique_name in unique_names:
            requested_name = clean_filename_column_name(unique_name)

            if not requested_name:
                continue

            output_column = get_safe_output_column_name(
                requested_name=requested_name,
                original_columns=original_column_set,
                generated_columns=generated_columns,
                source_value_column=value_column
            )

            if output_column is None:
                continue

            matching_rows = valid_name_mask & cleaned_names.eq(unique_name)

            if output_column not in generated_columns:
                generated_columns[output_column] = pd.Series(
                    pd.NA,
                    index=original_index,
                    dtype="object"
                )

            target_series = generated_columns[output_column]
            source_values = df[value_column]

            target_as_string = (
                target_series
                .astype("string")
                .str.strip()
                .str.lower()
            )

            destination_empty = (
                target_series.isna()
                | target_as_string.isna()
                | target_as_string.eq("")
                | target_as_string.eq("null")
                | target_as_string.eq("nan")
                | target_as_string.eq("none")
            )

            update_mask = matching_rows & destination_empty
            target_series.loc[update_mask] = source_values.loc[update_mask].values
            generated_columns[output_column] = target_series

        columns_to_drop.extend([name_column, value_column])

    columns_to_drop = list(dict.fromkeys(columns_to_drop))
    existing_columns_to_drop = [
        column for column in columns_to_drop if column in df.columns
    ]

    base_df = df.drop(columns=existing_columns_to_drop)

    if generated_columns:
        generated_df = pd.DataFrame(generated_columns, index=original_index)
        result_df = pd.concat([base_df, generated_df], axis=1)
    else:
        result_df = base_df

    result_df = result_df.copy()

    if len(result_df) != original_row_count:
        raise RuntimeError("Metadata expansion changed the number of rows.")

    if not result_df.index.equals(original_index):
        raise RuntimeError("Metadata expansion changed the DataFrame index.")

    return result_df


def drop_empty_columns(df):
    columns_to_drop = []

    for column in df.columns:
        cleaned_series = (
            df[column]
            .astype("string")
            .str.strip()
            .str.lower()
        )

        has_valid_data = (
            df[column].notna()
            & cleaned_series.notna()
            & cleaned_series.ne("")
            & cleaned_series.ne("null")
            & cleaned_series.ne("nan")
            & cleaned_series.ne("none")
        )

        if not has_valid_data.any():
            columns_to_drop.append(column)

    if columns_to_drop:
        df = df.drop(columns=columns_to_drop)

    return df


def rearrange_columns(df):
    columns = list(df.columns)
    front_columns = []

    if "rawEvent" in columns:
        columns.remove("rawEvent")
        front_columns.append("rawEvent")

    if "eventTime" in columns:
        columns.remove("eventTime")
        front_columns.append("eventTime")

    return df[front_columns + columns]


def load_input_file(filename):
    extension = os.path.splitext(filename)[1].lower()

    if extension == ".csv":
        return pd.read_csv(filename, low_memory=False)

    if extension == ".xlsx":
        return pd.read_excel(filename, engine="openpyxl")

    raise ValueError("Unsupported file type. Only CSV and XLSX are supported.")


def save_dataframe_to_excel(df, output_filename):
    df.to_excel(output_filename, index=False, engine="openpyxl")


def style_excel_file(output_filename):
    """Apply SOC Dark-theme formatting to the exported Excel file."""
    workbook = openpyxl.load_workbook(output_filename)
    sheet = workbook.active
    sheet.title = "Triggered Events"

    # Dark Border Styling
    thin_border = Border(
        left=Side(style="thin", color="334155"),
        right=Side(style="thin", color="334155"),
        top=Side(style="thin", color="334155"),
        bottom=Side(style="thin", color="334155")
    )

    # Header Styling (Dark Slate background with Cyan Font)
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")

    # Data Row Styling
    row_font = Font(name="Calibri", size=11, color="1E293B")
    row_alignment = Alignment(horizontal="center", vertical="center")

    for cell in sheet[1]:
        cell.border = thin_border
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment

    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.border = thin_border
            cell.font = row_font
            cell.alignment = row_alignment

    for column_cells in sheet.columns:
        max_length = 0
        column_letter = openpyxl.utils.get_column_letter(column_cells[0].column)
        column_header = column_cells[0].value

        for cell in column_cells:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))

        adjusted_width = (max_length + 2) * 1.2

        if column_header == "rawEvent":
            sheet.column_dimensions[column_letter].width = 50
        else:
            sheet.column_dimensions[column_letter].width = max(min(adjusted_width, 40), 12)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 25
    workbook.save(output_filename)


# ============================================================
# MAIN PROCESSING FUNCTION
# ============================================================

def update_progress(percent):
    root.after(
        0,
        lambda: (
            progress_value.set(percent),
            progress_label.config(text=f"{int(percent)}%")
        )
    )

def process_log_file(filename):
    total_start = time.time()

    set_stage("Loading input log file...")
    load_start = time.time()
    df = load_input_file(filename)
    load_seconds = time.time() - load_start
    update_progress(10)

    original_row_count = len(df)
    original_column_count = len(df.columns)

    write_log(f"[INFO] File loaded in {load_seconds:.2f} seconds.")
    write_log(f"[INFO] Input rows: {original_row_count}")
    write_log(f"[INFO] Input columns: {original_column_count}")

    set_stage("Rearranging columns...")
    df = rearrange_columns(df)
    update_progress(20)

    set_stage("Filtering out unwanted columns...")
    unwanted_columns = [
        "rawEventHash",
        "aisaacReceivedTime",
        "customerURI",
        "cenNifiReceiptTime",
        "logFilterKafkaInTime",
        "logFilterInTime"
    ]

    existing_unwanted_columns = [
        column for column in unwanted_columns if column in df.columns
    ]

    if existing_unwanted_columns:
        df = df.drop(columns=existing_unwanted_columns)
        write_log("[INFO] Removed columns: " + ", ".join(existing_unwanted_columns))
    else:
        write_log("[INFO] No predefined unwanted columns were found.")

    update_progress(35)

    set_stage("Expanding metadata fields...")
    metadata_start = time.time()
    df = expand_metadata_columns_fast(df)
    update_progress(55)
    metadata_seconds = time.time() - metadata_start
    write_log(f"[INFO] Metadata expanded in {metadata_seconds:.2f} seconds.")

    if len(df) != original_row_count:
        raise RuntimeError("Row count changed after metadata expansion.")

    set_stage("Purging empty data fields...")
    null_scan_start = time.time()
    df = drop_empty_columns(df)
    update_progress(70)
    null_scan_seconds = time.time() - null_scan_start
    write_log(f"[INFO] Empty columns dropped in {null_scan_seconds:.2f} seconds.")

    if len(df) != original_row_count:
        raise RuntimeError("Row count changed during processing.")

    set_stage('Normalizing blank values to "null"...')
    df = df.fillna("null")
    df.replace(
        to_replace=r"^\s*$",
        value="null",
        regex=True,
        inplace=True
    )
    update_progress(80)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    base_path, _ = os.path.splitext(filename)
    output_filename = f"{base_path}_Normalized_{timestamp}.xlsx"

    set_stage("Generating XLSX payload...")
    write_start = time.time()
    save_dataframe_to_excel(df, output_filename)
    update_progress(90)
    write_seconds = time.time() - write_start
    write_log(f"[INFO] Excel data written in {write_seconds:.2f} seconds.")

    set_stage("Applying SOC terminal formatting...")
    style_start = time.time()
    style_excel_file(output_filename)
    update_progress(100)
    style_seconds = time.time() - style_start
    write_log(f"[INFO] Styling complete in {style_seconds:.2f} seconds.")

    total_seconds = time.time() - total_start

    return {
        "output_filename": output_filename,
        "rows": len(df),
        "columns": len(df.columns),
        "elapsed": total_seconds
    }


# ============================================================
# GUI ACTIONS
# ============================================================

def select_file():
    global selected_filename

    if processing_active:
        messagebox.showwarning("Process Running", "A log normalization process is already active.")
        return

    filename = filedialog.askopenfilename(
        title="Select Security Log File",
        filetypes=[
            ("Supported Logs", "*.csv *.xlsx"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx"),
            ("All files", "*.*")
        ]
    )

    if not filename:
        return

    filename = clean_display_path(filename)
    extension = os.path.splitext(filename)[1].lower()

    if extension not in (".csv", ".xlsx"):
        messagebox.showerror("Invalid File Format", "Only CSV and XLSX files are accepted.")
        return

    if not os.path.isfile(filename):
        messagebox.showerror("File Error", "Selected log file could not be retrieved.")
        return

    selected_filename = filename
    selected_file_label.config(text=filename, fg=APP_TEXT_MAIN)
    start_processing(filename)


def start_processing(filename):
    global processing_active, processing_started_at

    processing_active = True
    processing_started_at = time.time()

    select_button.config(state=tk.DISABLED, bg=APP_CARD_BG)
    open_file_button.config(state=tk.DISABLED, bg=APP_CARD_BG)
    progress_value.set(0)
    progress_label.config(text="0%")

    set_status("INITIALIZING LOG NORMALIZATION...", APP_ACCENT)
    write_log(f"[INFO] Processing Target: {filename}")

    worker = threading.Thread(
        target=processing_worker,
        args=(filename,),
        daemon=True
    )
    worker.start()
    root.after(1000, monitor_processing_time)


def processing_worker(filename):
    try:
        result = process_log_file(filename)
        root.after(0, processing_completed, result)

    except PermissionError:
        error_message = "File Write Access Denied. Close the open spreadsheet and retry."
        root.after(0, processing_failed, error_message)

    except MemoryError:
        error_message = "System Out of Memory. Resource exhaustion prevents parsing."
        root.after(0, processing_failed, error_message)

    except Exception as error:
        root.after(0, processing_failed, str(error))


def processing_completed(result):
    global processing_active, last_output_filename

    processing_active = False
    last_output_filename = result["output_filename"]

    progress_value.set(100)
    progress_label.config(text="100%")
    select_button.config(state=tk.NORMAL, bg=APP_BTN_PRIMARY, fg="white")
    
    # Enabled state with active blue background and white text
    open_file_button.config(state=tk.NORMAL, bg=APP_BTN_PRIMARY, fg="white")

    set_status("EXECUTION SUCCESS", APP_SUCCESS)

    messagebox.showinfo(
        "Operation Complete",
        "Log Normalization Finished Successfully.\n\n"
        f"Rows Parsed: {result['rows']}\n"
        f"Output Fields: {result['columns']}\n"
        f"Execution Time: {result['elapsed']:.2f}s\n\n"
        f"Output Target:\n{result['output_filename']}"
    )


def processing_failed(error_message):
    global processing_active

    processing_active = False
    select_button.config(state=tk.NORMAL, bg=APP_BTN_PRIMARY)

    set_status("EXECUTION ABORTED // ERROR ENCOUNTERED", APP_ERROR)
    messagebox.showerror("Normalization Failure", error_message)


def monitor_processing_time():
    if not processing_active or processing_started_at is None:
        return

    elapsed = time.time() - processing_started_at

    if elapsed >= 60:
        status_label.config(
            text="PROCESSING LARGE LOG VOLUME...",
            fg=APP_WARNING
        )

    root.after(1000, monitor_processing_time)


def open_output_file():
    if not last_output_filename:
        return

    try:
        if sys.platform.startswith("win"):
            os.startfile(last_output_filename)
        elif sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", last_output_filename])
        else:
            import subprocess
            subprocess.Popen(["xdg-open", last_output_filename])

    except Exception as error:
        messagebox.showerror("IO Error", f"Unable to access output file.\n\n{error}")

def exit_application():
    if processing_active:
        should_exit = messagebox.askyesno(
            "Abort Execution",
            "A log normalization routine is active.\n\nAbort process and terminate interface?"
        )
        if not should_exit:
            return

    root.destroy()


# ============================================================
# UI CONSTRUCTION (SOC DESIGN SYSTEM)
# ============================================================

root = tk.Tk()
root.title(APP_TITLE)
root.geometry("560x380")
root.resizable(False, False)
root.configure(bg=APP_BACKGROUND)
progress_value = tk.DoubleVar(value=0)
root.protocol("WM_DELETE_WINDOW", exit_application)

# ttk Style Customization for Progressbar
style = ttk.Style()
style.theme_use("default")
style.configure(
    "Custom.Horizontal.TProgressbar",
    thickness=8,
    troughcolor=APP_CARD_BG,
    background=APP_ACCENT,
    bordercolor=APP_BACKGROUND,
    borderwidth=0
)

main_frame = tk.Frame(root, bg=APP_BACKGROUND)
main_frame.pack(fill=tk.BOTH, expand=True, padx=24, pady=20)

# Header Banner
title_label = tk.Label(
    main_frame,
    text="🛡️ FORMATIX",
    bg=APP_CARD_BG,
    fg=APP_ACCENT,
    font=("Consolas", 16, "bold"),
    pady=12,
    relief=tk.FLAT
)
title_label.pack(fill=tk.X)

# Operational Description
instruction_label = tk.Label(
    main_frame,
    text="Select a CSV or XLSX file to normalize and export it as a formatted Excel file.",
    bg=APP_BACKGROUND,
    fg=APP_TEXT_MUTED,
    font=("Segoe UI", 9),
    wraplength=520,
    justify=tk.CENTER
)
instruction_label.pack(pady=(15, 12))

# Selection Container Card
card_frame = tk.Frame(main_frame, bg=APP_CARD_BG, padx=15, pady=12)
card_frame.pack(fill=tk.X, pady=(0, 15))

selected_file_title = tk.Label(
    card_frame,
    text="TARGET LOG SOURCE:",
    bg=APP_CARD_BG,
    fg=APP_TEXT_MUTED,
    font=("Consolas", 9, "bold")
)
selected_file_title.pack(anchor="w", pady=(0, 4))

selected_file_label = tk.Label(
    card_frame,
    text="No log file loaded...",
    bg="#0F172A",
    fg="#475569",
    font=("Consolas", 9),
    relief=tk.FLAT,
    anchor="w",
    padx=10,
    pady=8,
    wraplength=480,
    justify=tk.LEFT
)
selected_file_label.pack(fill=tk.X)

# Button Controls
button_frame = tk.Frame(main_frame, bg=APP_BACKGROUND)
button_frame.pack(pady=(0, 15))

select_button = tk.Button(
    button_frame,
    text="Select Log File",
    command=select_file,
    bg=APP_BTN_PRIMARY,
    fg="white",
    activebackground=APP_BTN_HOVER,
    activeforeground="white",
    font=("Segoe UI", 10, "bold"),
    width=15,
    relief=tk.FLAT,
    cursor="hand2"
)
select_button.grid(row=0, column=0, padx=6)

open_file_button = tk.Button(
    button_frame,
    text="Open Export File",
    command=open_output_file,
    bg=APP_CARD_BG,
    fg=APP_TEXT_MUTED,
    activebackground=APP_BTN_HOVER,
    activeforeground="white",
    disabledforeground="#475569",  # Dimmed text when disabled
    font=("Segoe UI", 10, "bold"),  # Bold font weight matching Select button
    width=15,
    state=tk.DISABLED,
    relief=tk.FLAT,
    cursor="hand2"
)
open_file_button.grid(row=0, column=1, padx=6)

exit_button = tk.Button(
    button_frame,
    text="Exit System",
    command=exit_application,
    bg=APP_CARD_BG,
    fg=APP_TEXT_MUTED,
    activebackground=APP_ERROR,
    activeforeground="white",
    font=("Segoe UI", 10),
    width=12,
    relief=tk.FLAT,
    cursor="hand2"
)
exit_button.grid(row=0, column=2, padx=6)

# Progress Diagnostics
progress_bar = ttk.Progressbar(
    main_frame,
    variable=progress_value,
    maximum=100,
    mode="determinate",
    style="Custom.Horizontal.TProgressbar"
)
progress_bar.pack(fill=tk.X, pady=(0, 4))

progress_label = tk.Label(
    main_frame,
    text="0%",
    bg=APP_BACKGROUND,
    fg=APP_TEXT_MUTED,
    font=("Consolas", 8, "bold")
)
progress_label.pack()

status_label = tk.Label(
    main_frame,
    text="SYSTEM READY",
    bg=APP_BACKGROUND,
    fg=APP_SUCCESS,
    font=("Consolas", 10, "bold")
)
status_label.pack(pady=(2, 6))

developer_label = tk.Label(
    main_frame,
    text="DEV //Gourab Sarkar",
    bg=APP_BACKGROUND,
    fg="#475569",
    font=("Consolas", 8)
)
developer_label.pack(anchor="e")


if __name__ == "__main__":
    root.mainloop()