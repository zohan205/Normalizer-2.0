import os
import sys
import time
import threading
from datetime import datetime

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import tkinter as tk
from tkinter import filedialog, messagebox




# ============================================================
# APPLICATION STATE
# ============================================================

APP_TITLE = "Formatix"
APP_BACKGROUND = "#06111F"
APP_PANEL = "#091A2D"
APP_PANEL_2 = "#0D223A"
APP_ACCENT = "#129BFF"
APP_ACCENT_BRIGHT = "#00D9FF"
APP_TEXT = "#F4F8FF"
APP_MUTED = "#91A7C0"
APP_SUCCESS = "#2DFF88"
APP_ERROR = "#FF5C7A"
APP_WARNING = "#FFB547"
APP_BORDER = "#124A78"

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

def set_status(message, color="#333333"):
    """Update the status label safely from any thread."""
    root.after(0, lambda: status_label.config(text=message, fg=color))


def set_stage(message):
    """Update both the status label and activity log."""
    set_status(message, "#333333")
    write_log(message)


def clean_display_path(path):
    """Return a readable absolute path."""
    return os.path.abspath(os.path.expanduser(str(path)))


# ============================================================
# DATA CLEANING FUNCTIONS
# ============================================================

def clean_name_series(series):
    """
    Clean a metadata Name column.

    Invalid values:
    - Empty
    - null
    - nan
    - none
    """
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
    """
    Clean a dynamic Excel column name.

    Excel column headers cannot be empty. Line breaks and tab
    characters are replaced with spaces.
    """
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
    """Return a safe output column name without overwriting data."""
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
    """
    Convert metadata Name/value pairs into separate columns.

    Important:
    - No rows are added.
    - No rows are deleted.
    - Row order is not changed.
    - Non-applicable cells remain empty.
    - New columns are added once using pd.concat().
    """
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
    """Remove columns containing only empty/null/nan/none values."""
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
    """Place rawEvent first and eventTime second."""
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
    """Load a CSV or XLSX file."""
    extension = os.path.splitext(filename)[1].lower()

    if extension == ".csv":
        return pd.read_csv(filename, low_memory=False)

    if extension == ".xlsx":
        return pd.read_excel(filename, engine="openpyxl")

    raise ValueError("Unsupported file type. Only CSV and XLSX are supported.")


def save_dataframe_to_excel(df, output_filename):
    """Save the cleaned DataFrame to Excel."""
    df.to_excel(output_filename, index=False, engine="openpyxl")


def style_excel_file(output_filename):
    """Apply formatting to the generated Excel workbook."""
    workbook = openpyxl.load_workbook(output_filename)
    sheet = workbook.active
    sheet.title = "Triggered Events"

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    header_font = Font(
        name="Calibri",
        size=13,
        bold=True,
        color="FFFFFF"
    )

    header_fill = PatternFill(
        start_color="4F81BD",
        end_color="4F81BD",
        fill_type="solid"
    )

    header_alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    row_font = Font(
        name="Calibri",
        size=12,
        color="000000"
    )

    row_alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

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
            sheet.column_dimensions[column_letter].width = max(
                min(adjusted_width, 40),
                10
            )

    sheet.freeze_panes = "A2"

    sheet.auto_filter.ref = None
    sheet.row_dimensions[1].height = 22
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
    """Normalize the selected file and return processing details."""
    total_start = time.time()

    set_stage("Loading input file...")
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

    set_stage("Removing predefined unwanted columns...")
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
        write_log(
            "[INFO] Removed columns: "
            + ", ".join(existing_unwanted_columns)
        )
    else:
        write_log("[INFO] No predefined unwanted columns were found.")

    update_progress(35)

    set_stage("Expanding metadata columns...")
    metadata_start = time.time()
    df = expand_metadata_columns_fast(df)
    update_progress(55)
    metadata_seconds = time.time() - metadata_start
    write_log(
        f"[INFO] Metadata columns expanded in {metadata_seconds:.2f} seconds."
    )

    if len(df) != original_row_count:
        raise RuntimeError("Row count changed after metadata expansion.")

    set_stage("Removing empty columns...")
    null_scan_start = time.time()
    df = drop_empty_columns(df)
    update_progress(70)
    null_scan_seconds = time.time() - null_scan_start
    write_log(
        f"[INFO] Empty columns removed in {null_scan_seconds:.2f} seconds."
    )

    if len(df) != original_row_count:
        raise RuntimeError("Row count changed during processing.")

    set_stage('Converting blank values to "null"...')
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
    output_filename = f"{base_path}_{timestamp}.xlsx"

    set_stage("Writing Excel file...")
    write_start = time.time()
    save_dataframe_to_excel(df, output_filename)
    update_progress(90)
    write_seconds = time.time() - write_start
    write_log(f"[INFO] Excel data written in {write_seconds:.2f} seconds.")

    set_stage("Applying Excel formatting...")
    style_start = time.time()
    style_excel_file(output_filename)
    update_progress(100)
    style_seconds = time.time() - style_start
    write_log(
        f"[INFO] Excel styling completed in {style_seconds:.2f} seconds."
    )

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
    """Open the file picker and start processing the selected file."""
    global selected_filename

    if processing_active:
        messagebox.showwarning(
            "Processing in progress",
            "A file is already being processed."
        )
        return

    filename = filedialog.askopenfilename(
        title="Select CSV or XLSX File",
        filetypes=[
            ("Supported files", "*.csv *.xlsx"),
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
        messagebox.showerror(
            "Unsupported file",
            "Only CSV and XLSX files are supported."
        )
        return

    if not os.path.isfile(filename):
        messagebox.showerror("File not found", "The selected file was not found.")
        return

    selected_filename = filename
    selected_file_label.config(text=filename)


    start_processing(filename)


def start_processing(filename):
    """Start file processing on a background thread."""
    global processing_active, processing_started_at

    processing_active = True
    processing_started_at = time.time()

    select_button.config(state=tk.DISABLED)
    open_file_button.config(state=tk.DISABLED)
    progress_value.set(0)
    progress_label.config(text="0%")

    set_status("Processing...", "#333333")
    write_log(f"[INFO] Selected file: {filename}")

    worker = threading.Thread(
        target=processing_worker,
        args=(filename,),
        daemon=True
    )
    worker.start()

    root.after(1000, monitor_processing_time)


def processing_worker(filename):
    """Run data processing without freezing the GUI."""
    try:
        result = process_log_file(filename)
        root.after(0, processing_completed, result)

    except PermissionError:
        error_message = (
            "Permission denied while writing the output file.\n\n"
            "Close the output Excel file if it is already open, "
            "then try again."
        )
        root.after(0, processing_failed, error_message)

    except MemoryError:
        error_message = (
            "The computer does not have enough available memory "
            "to process this file."
        )
        root.after(0, processing_failed, error_message)

    except Exception as error:
        root.after(0, processing_failed, str(error))


def processing_completed(result):
    """Handle successful processing on the main GUI thread."""
    global processing_active, last_output_filename

    processing_active = False
    last_output_filename = result["output_filename"]

    progress_value.set(100)
    progress_label.config(text="100%")
    select_button.config(state=tk.NORMAL)
    open_file_button.config(state=tk.NORMAL)

    set_status("Processing completed successfully", APP_SUCCESS)

    write_log("")
    write_log("[SUCCESS] Processing completed.")
    write_log(f"[SUCCESS] Rows preserved: {result['rows']}")
    write_log(f"[SUCCESS] Final columns: {result['columns']}")
    write_log(f"[SUCCESS] Total elapsed time: {result['elapsed']:.2f} seconds")
    write_log(f"[SUCCESS] Saved as: {result['output_filename']}")

    messagebox.showinfo(
        "Processing completed",
        "The file was processed successfully.\n\n"
        f"Rows preserved: {result['rows']}\n"
        f"Final columns: {result['columns']}\n"
        f"Elapsed time: {result['elapsed']:.2f} seconds\n\n"
        f"Saved as:\n{result['output_filename']}"
    )


def processing_failed(error_message):
    """Handle processing failure on the main GUI thread."""
    global processing_active

    processing_active = False
    progress_bar.stop()
    select_button.config(state=tk.NORMAL)

    set_status("Processing failed", APP_ERROR)
    write_log(f"[ERROR] {error_message}")

    messagebox.showerror("Processing failed", error_message)


def monitor_processing_time():
    """Show a notice if processing has run for more than 60 seconds."""
    if not processing_active or processing_started_at is None:
        return

    elapsed = time.time() - processing_started_at

    if elapsed >= 60:
        status_label.config(
            text="Still processing. Large files can take longer.",
            fg=APP_WARNING
        )

    root.after(1000, monitor_processing_time)


def open_output_file():
    """Open the generated Excel file."""
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
        messagebox.showerror(
            "Unable to open file",
            f"Could not open the output file.\n\n{error}"
        )

def exit_application():
    """Close the program, with a warning during active processing."""
    if processing_active:
        should_exit = messagebox.askyesno(
            "Exit Formatix",
            "A file is currently being processed.\n\n"
            "Exiting now will stop the operation. Do you want to exit?"
        )

        if not should_exit:
            return

    root.destroy()


# ============================================================
# UI CONSTRUCTION - SOC / CYBER SECURITY THEME
# ============================================================

from tkinter import ttk

root = tk.Tk()
root.title(APP_TITLE)
root.geometry("1180x760")
root.resizable(False, False)
root.configure(bg=APP_BACKGROUND)
progress_value = tk.DoubleVar(value=0)
root.protocol("WM_DELETE_WINDOW", exit_application)

# ---------- Theme ----------
FONT = "Segoe UI"
MONO = "Consolas"

style = ttk.Style(root)
try:
    style.theme_use("clam")
except tk.TclError:
    pass

style.configure(
    "SOC.Horizontal.TProgressbar",
    troughcolor="#0B1D31",
    background=APP_ACCENT,
    bordercolor=APP_BORDER,
    lightcolor=APP_ACCENT,
    darkcolor=APP_ACCENT,
    thickness=10,
)

# ---------- Background canvas ----------
background_canvas = tk.Canvas(
    root,
    bg=APP_BACKGROUND,
    highlightthickness=0,
    bd=0,
)
background_canvas.place(x=0, y=0, relwidth=1, relheight=1)


def draw_cyber_background():
    """Draw subtle SOC/network circuitry without external image assets."""
    w = root.winfo_width() or 1180
    h = root.winfo_height() or 760

    # Faint grid
    for x in range(20, w, 55):
        background_canvas.create_line(x, 0, x, h, fill="#081A2D", width=1)
    for y in range(20, h, 55):
        background_canvas.create_line(0, y, w, y, fill="#081A2D", width=1)

    # Network nodes / connections on the right
    nodes = [(900, 110), (1010, 70), (1100, 135), (1160, 80),
             (970, 190), (1080, 220), (1150, 300), (1030, 310)]
    edges = [(0, 1), (1, 2), (2, 3), (0, 4), (4, 5), (5, 6),
             (4, 7), (7, 5), (2, 5)]
    for a, b in edges:
        x1, y1 = nodes[a]
        x2, y2 = nodes[b]
        background_canvas.create_line(
            x1, y1, x2, y2,
            fill="#0C3558", width=1
        )
    for x, y in nodes:
        background_canvas.create_oval(
            x - 3, y - 3, x + 3, y + 3,
            fill=APP_ACCENT, outline=""
        )

    # Decorative scan line
    background_canvas.create_line(
        0, 94, w, 94, fill="#0C3558", width=1
    )


# ---------- Header ----------
header = tk.Frame(root, bg="#07182A", height=94)
header.place(x=0, y=0, relwidth=1)
header.pack_propagate(False)

# Shield icon drawn with a Canvas
shield = tk.Canvas(header, width=58, height=66, bg="#07182A", highlightthickness=0)
shield.place(x=30, y=14)
shield.create_polygon(
    29, 3, 51, 11, 49, 38, 29, 60, 9, 38, 7, 11,
    fill="#0A75D8", outline="#00D9FF", width=2
)
shield.create_polygon(
    29, 10, 44, 16, 42, 35, 29, 50, 16, 35, 14, 16,
    fill="#09223A", outline=""
)
shield.create_rectangle(23, 27, 35, 39, fill="#DFF8FF", outline="")
shield.create_arc(21, 18, 37, 34, start=0, extent=180, style=tk.ARC,
                  outline="#DFF8FF", width=3)

brand = tk.Label(
    header, text="Formatix", bg="#07182A", fg=APP_TEXT,
    font=(FONT, 25, "bold")
)
brand.place(x=105, y=18)

# Accent part of the product name
tk.Label(
    header, text="SECURE  /  ANALYZE  /  EXPORT",
    bg="#07182A", fg=APP_MUTED, font=(MONO, 9, "bold")
).place(x=107, y=58)

# Header telemetry
telemetry = tk.Frame(header, bg="#07182A")
telemetry.place(relx=0.68, y=20, relwidth=0.29, height=55)

tk.Label(
    telemetry, text="SOC DATA PIPELINE",
    bg="#07182A", fg=APP_ACCENT_BRIGHT,
    font=(MONO, 9, "bold")
).pack(anchor="e")

tk.Label(
    telemetry, text="NORMALIZE  •  VALIDATE  •  EXPORT",
    bg="#07182A", fg=APP_MUTED,
    font=(FONT, 9)
).pack(anchor="e", pady=(4, 0))

# ---------- Main content ----------
main_frame = tk.Frame(root, bg=APP_BACKGROUND)
main_frame.place(x=35, y=118, width=1110, height=580)

# Intro / mission area
intro = tk.Frame(main_frame, bg=APP_BACKGROUND)
intro.pack(fill=tk.X, pady=(0, 18))

tk.Label(
    intro,
    text="SOC EVENT FORMATTER",
    bg=APP_BACKGROUND,
    fg=APP_ACCENT_BRIGHT,
    font=(MONO, 10, "bold")
).pack(anchor="w")

tk.Label(
    intro,
    text="Normalize security telemetry into a clean, analyst-ready Excel file.",
    bg=APP_BACKGROUND,
    fg=APP_TEXT,
    font=(FONT, 18, "bold")
).pack(anchor="w", pady=(5, 2))

tk.Label(
    intro,
    text="Select a CSV or XLSX source. Formatix preserves rows, expands metadata, removes noise, and applies SOC-friendly Excel formatting.",
    bg=APP_BACKGROUND,
    fg=APP_MUTED,
    font=(FONT, 9),
    wraplength=920,
    justify=tk.LEFT
).pack(anchor="w")

# ---------- Main processing card ----------
card = tk.Frame(
    main_frame,
    bg=APP_PANEL,
    highlightbackground=APP_BORDER,
    highlightcolor=APP_ACCENT,
    highlightthickness=1,
)
card.pack(fill=tk.X)

card_inner = tk.Frame(card, bg=APP_PANEL)
card_inner.pack(fill=tk.X, padx=28, pady=25)

# Section heading
heading_row = tk.Frame(card_inner, bg=APP_PANEL)
heading_row.pack(fill=tk.X)

file_icon = tk.Canvas(heading_row, width=48, height=54, bg=APP_PANEL, highlightthickness=0)
file_icon.pack(side=tk.LEFT, padx=(0, 14))
file_icon.create_rectangle(9, 4, 38, 47, outline=APP_ACCENT, width=2)
file_icon.create_polygon(29, 4, 38, 13, 29, 13, fill=APP_ACCENT, outline=APP_ACCENT)
file_icon.create_text(23, 31, text="</>", fill=APP_ACCENT_BRIGHT, font=(MONO, 9, "bold"))

tk.Label(
    heading_row,
    text="SECURITY TELEMETRY INPUT",
    bg=APP_PANEL,
    fg=APP_TEXT,
    font=(FONT, 12, "bold")
).pack(anchor="w", pady=(4, 0))

tk.Label(
    heading_row,
    text="Supported sources: CSV / XLSX",
    bg=APP_PANEL,
    fg=APP_MUTED,
    font=(FONT, 9)
).pack(anchor="w", pady=(3, 0))

# File selector box
file_box = tk.Frame(
    card_inner,
    bg="#07182A",
    highlightbackground="#14547F",
    highlightthickness=1,
)
file_box.pack(fill=tk.X, pady=(20, 18))

selected_file_title = tk.Label(
    file_box,
    text="  SELECTED FILE",
    bg="#07182A",
    fg=APP_ACCENT_BRIGHT,
    font=(MONO, 8, "bold")
)
selected_file_title.pack(anchor="w", padx=12, pady=(9, 3))

selected_file_label = tk.Label(
    file_box,
    text="No file selected",
    bg="#07182A",
    fg="#A9C4E0",
    font=(MONO, 10),
    anchor="w",
    padx=12,
    pady=10,
)
selected_file_label.pack(fill=tk.X, side=tk.LEFT, expand=True)

# Folder icon/button on the right
folder_button = tk.Button(
    file_box,
    text="▣",
    command=select_file,
    bg="#0A2A46",
    fg=APP_ACCENT_BRIGHT,
    activebackground="#0F3C62",
    activeforeground="white",
    bd=0,
    font=(FONT, 18, "bold"),
    cursor="hand2",
    width=4,
)
folder_button.pack(side=tk.RIGHT, padx=6, pady=7)

# ---------- Buttons ----------
button_frame = tk.Frame(card_inner, bg=APP_PANEL)
button_frame.pack(fill=tk.X, pady=(0, 20))


def make_button(parent, text, command, width, enabled=True):
    btn = tk.Button(
        parent,
        text=text,
        command=command,
        bg="#0E3150" if enabled else "#13243A",
        fg=APP_TEXT if enabled else "#5E7690",
        activebackground="#1268A5",
        activeforeground="white",
        disabledforeground="#536B83",
        relief=tk.FLAT,
        bd=0,
        highlightbackground="#28658F",
        highlightthickness=1,
        font=(FONT, 10, "bold"),
        width=width,
        height=2,
        cursor="hand2" if enabled else "arrow",
    )
    return btn


def bind_hover(btn, normal="#0E3150", hover="#1268A5"):
    btn.bind("<Enter>", lambda e: btn.config(bg=hover) if str(btn["state"]) != "disabled" else None)
    btn.bind("<Leave>", lambda e: btn.config(bg=normal) if str(btn["state"]) != "disabled" else None)

select_button = make_button(button_frame, "▣   SELECT FILE", select_file, 22)
select_button.pack(side=tk.LEFT, padx=(0, 12))
bind_hover(select_button)

open_file_button = make_button(button_frame, "↥   OPEN OUTPUT FILE", open_output_file, 25, enabled=False)
open_file_button.config(state=tk.DISABLED)
open_file_button.pack(side=tk.LEFT, padx=(0, 12))

exit_button = make_button(button_frame, "⇥   EXIT", exit_application, 14)
exit_button.pack(side=tk.LEFT)
bind_hover(exit_button, normal="#0E3150", hover="#8B2941")

# ---------- Progress / operation telemetry ----------
progress_panel = tk.Frame(
    card_inner,
    bg="#07182A",
    highlightbackground="#14547F",
    highlightthickness=1,
)
progress_panel.pack(fill=tk.X)

progress_top = tk.Frame(progress_panel, bg="#07182A")
progress_top.pack(fill=tk.X, padx=16, pady=(12, 4))

tk.Label(
    progress_top,
    text="PROCESSING PIPELINE",
    bg="#07182A",
    fg=APP_MUTED,
    font=(MONO, 8, "bold")
).pack(side=tk.LEFT)

progress_label = tk.Label(
    progress_top,
    text="0%",
    bg="#07182A",
    fg=APP_ACCENT_BRIGHT,
    font=(MONO, 9, "bold")
)
progress_label.pack(side=tk.RIGHT)

progress_bar = ttk.Progressbar(
    progress_panel,
    variable=progress_value,
    maximum=100,
    mode="determinate",
    style="SOC.Horizontal.TProgressbar",
)
progress_bar.pack(fill=tk.X, padx=16, pady=(0, 11))

status_row = tk.Frame(progress_panel, bg="#07182A")
status_row.pack(fill=tk.X, padx=16, pady=(0, 11))

status_dot = tk.Canvas(status_row, width=12, height=12, bg="#07182A", highlightthickness=0)
status_dot.pack(side=tk.LEFT, padx=(0, 7))
status_dot.create_oval(2, 2, 10, 10, fill=APP_SUCCESS, outline="")

status_label = tk.Label(
    status_row,
    text="Ready",
    bg="#07182A",
    fg=APP_SUCCESS,
    font=(FONT, 10, "bold")
)
status_label.pack(side=tk.LEFT)

# ---------- Footer SOC strip ----------
footer = tk.Frame(root, bg="#07182A", height=62)
footer.place(x=0, y=698, relwidth=1)
footer.pack_propagate(False)


def footer_item(parent, icon, title, x, width):
    frame = tk.Frame(parent, bg="#07182A", width=width, height=62)
    frame.place(x=x, y=0)
    frame.pack_propagate(False)
    tk.Label(frame, text=icon, bg="#07182A", fg=APP_ACCENT,
             font=(FONT, 16, "bold")).pack(side=tk.LEFT, padx=(8, 8))
    tk.Label(frame, text=title, bg="#07182A", fg=APP_MUTED,
             font=(MONO, 8, "bold")).pack(side=tk.LEFT)
    return frame

footer_item(footer, "◈", "SOC TOOLKIT", 30, 220)
footer_item(footer, "⌕", "INVESTIGATE", 300, 180)
footer_item(footer, "◇", "PROTECT", 520, 160)
footer_item(footer, "↗", "RESPOND", 710, 160)

# Right-side operational label
tk.Label(
    footer,
    text="CYBER SECURITY\nOPERATIONS",
    bg="#07182A",
    fg=APP_ACCENT_BRIGHT,
    font=(MONO, 8, "bold"),
    justify=tk.LEFT,
).place(relx=0.86, y=13)

# Small designer credit
# Kept subtle so it doesn't compete with the SOC interface.
tk.Label(
    root,
    text="FORMATIX  •  SOC DATA NORMALIZATION ENGINE",
    bg=APP_BACKGROUND,
    fg="#47627D",
    font=(MONO, 7)
).place(x=38, y=682)

# Draw decorative background after geometry has been established.
root.after(100, draw_cyber_background)


if __name__ == "__main__":
    root.mainloop()
