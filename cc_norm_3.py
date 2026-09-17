import os
import sys
import time
import re
import threading
import ipaddress
from datetime import datetime

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# ============================================================
# APPLICATION STATE (SOC DARK THEME)
# ============================================================

APP_TITLE = "cc_norm // SOC Log Normalizer"
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
current_df = None
current_customer = None
current_vlans = []

ANALYSIS_COLUMNS = ["productHostname", "srcIPAddress", "dstIPAddress", "requestURL"]


# ============================================================
# GUI-SAFE HELPERS & FILENAME FORMATTERS
# ============================================================

def set_status(message, color=APP_TEXT_MUTED):
    """Update the status label safely from any thread."""
    root.after(0, lambda: status_label.config(text=message, fg=color))


def set_stage(message):
    """Update status label stage description."""
    set_status(message, APP_TEXT_MUTED)

def set_vlan_debug(message):
    try:
        vlan_debug_label.config(text=message)
    except:
        pass


def clean_display_path(path):
    """Return a readable absolute path."""
    return os.path.abspath(os.path.expanduser(str(path)))


def generate_formatted_filename(filename, ticket_id, segregate_col=None):
    """
    Generates output filename formatted as:
    Single Sheet: {Prefix}-{Ticket_ID}-{DD-Mon-YYYY}.xlsx
    Multi-Sheet:  {Prefix}-{Ticket_ID}-{DD-Mon-YYYY}_{fieldName}.xlsx
    """
    base_name = os.path.splitext(os.path.basename(filename))[0]
    directory = os.path.dirname(filename)

    # Extract leading prefix (e.g., 'PLOI-LM_Events-17-Sep-2026_1238' -> 'PLO')
    match = re.match(r"^([A-Za-z]+)", base_name)
    prefix = match.group(1) if match else "LOG"

    # Extract current date formatted as DD-Mon-YYYY (e.g. 17-Sep-2026)
    date_str = datetime.now().strftime("%d-%b-%Y")

    if segregate_col:
        out_name = f"{prefix}-{ticket_id}-{date_str}_{segregate_col}.xlsx"
    else:
        out_name = f"{prefix}-{ticket_id}-{date_str}.xlsx"

    return os.path.join(directory, out_name)


# ============================================================
# CUSTOMER / VLAN IDENTIFICATION
# ============================================================
def _log_sample(df, filename, rows=75):
    parts = [os.path.basename(filename), " ".join(map(str, df.columns))]
    if not df.empty:
        parts.append(" ".join(df.head(rows).astype(str).fillna("").values.ravel()))
    return " ".join(parts).lower()


def detect_customer(df, filename):

    base_name = os.path.basename(filename).upper()

    # Check filename first
    if base_name.startswith("RBIK"):
        return "RBIK"

    if base_name.startswith("Dunki"):
        return "Dunki"

    if base_name.startswith("bobla"):
        return "bobla obb"

    if base_name.startswith("Bees"):
        return "Bees"

    # Fallback to content search if filename doesn't match
    sample = _log_sample(df, filename)

    if "RBIK" in sample:
        return "RBIK"

    if "Dunki" in sample:
        return "Dunki"

    if "bobla" in sample:
        return "bobla obb"

    if "Bees" in sample:
        return "Bees"

    return None


def _extract_private_ips(df):
    """
    Extract only PRIVATE source/destination IPs.
    """
    ip_columns = []

    for col in df.columns:
        if str(col).lower() in [
            "srcipaddress",
            "dstipaddress"
        ]:
            ip_columns.append(col)

    if not ip_columns:
        return []

    private_ips = set()

    for col in ip_columns:
        for value in df[col].astype(str):

            try:
                ip = ipaddress.ip_address(value.strip())

                if ip.is_private:
                    private_ips.add(str(ip))

            except ValueError:
                pass

    return sorted(private_ips)


def _in(ip, cidr):
    return ipaddress.ip_address(ip) in ipaddress.ip_network(cidr, strict=False)

def detect_vlans(df, filename):

    customer = detect_customer(df, filename)

    debug_lines = []

    debug_lines.append(
        f"FILE : {os.path.basename(filename)}"
    )

    debug_lines.append(
        f"CUSTOMER : {customer}"
    )
    if not customer: return None, []
    sample, ips, vlans = _log_sample(df, filename), _extract_private_ips(df), set()
    debug_lines.append("")
    debug_lines.append(f"PRIVATE IPS FOUND : {len(ips)}")

    for ip in ips[:20]:
        debug_lines.append(f"IP : {ip}")

    if customer in ("Dunki", "bobla obb"):
        known = {10,20,30,40,50,60,70,72,80,90}
        for ip in ips:
            if _in(ip,"10.128.0.0/9"): vlans.add("10")
            elif _in(ip,"192.168.0.0/24"): vlans.add("20")
            elif _in(ip,"192.168.1.0/24"): vlans.add("40")
            elif customer == "bobla obb" and _in(ip,"10.10.10.0/24"): vlans.add("20")
            elif ip.startswith("192.168."):
                third=int(ip.split('.')[2])
                if third in known: vlans.add(str(third))
    elif customer == "RBIK":
        plk=any(x in sample for x in ("popeyes","plk")); th=any(x in sample for x in ("tim hortons","timhortons"))
        cbk="carrols bk" in sample; cplk="carrols plk" in sample
        rules=[]
        if cplk: rules += [(15,"192.168.1.0/24"),(10,"10.161.0.0/27"),(20,"10.116.0.0/27"),(30,"10.163.0.0/27"),(40,"10.164.0.0/27"),(50,"10.166.0.0/27"),(70,"10.168.0.0/27")]
        elif cbk: rules += [(10,"10.115.0.0/27"),(20,"10.116.0.0/27"),(30,"10.117.0.0/27"),(40,"10.118.0.0/27"),(50,"10.120.0.0/27"),(60,"10.121.0.0/28"),(70,"10.122.0.0/27"),(100,"10.114.0.0/27")]
        elif plk: rules += [(10,"192.168.16.0/24"),(20,"192.168.170.0/24")]
        elif th: rules += [(10,"192.168.11.0/24"),(20,"192.168.2.64/26"),(30,"192.168.2.208/28"),(50,"192.168.50.0/29"),(70,"192.168.1.0/24"),(99,"192.168.99.0/29"),(100,"192.168.100.0/23")]
        rules += [(10,"192.168.80.0/24"),(10,"192.168.85.0/24"),(10,"192.168.1.0/24"),(20,"192.168.2.0/24"),(30,"192.168.3.0/24"),(40,"192.168.4.0/24"),(50,"192.168.5.0/24"),(70,"192.168.7.0/24"),(90,"192.168.9.0/24"),(99,"192.168.99.0/24"),(200,"192.168.200.0/24"),(254,"192.168.254.0/24"),(60,"10.121.0.0/28"),(60,"10.167.0.0/24"),(100,"10.160.0.0/27")]
        for ip in ips:
            for vlan,cidr in rules:
                if _in(ip,cidr): vlans.add(str(vlan)); break
    else:
        rules=[("POS","10.60.0.0/16"),("fortilink","10.65.0.0/15"),("AccessPoint_Mgt","10.67.0.0/15"),("nac_segment","10.255.11.0/24"),("GUEST","10.0.0.0/16"),("Private","192.168.20.0/24"),("DVR","192.168.30.0/24"),("GUEST","192.168.40.0/24"),("AI","192.168.50.0/24"),("VoIP","192.168.60.0/24"),("quarantine","169.254.11.0/24"),("voice","169.254.12.0/24"),("video","169.254.13.0/24"),("onboarding","169.254.15.0/24"),("_default","169.254.3.0/24")]
        for ip in ips:
            for vlan,cidr in rules:
                if _in(ip,cidr): vlans.add(vlan); break
    key=lambda v:(0,int(v)) if v.isdigit() else (1,v.lower())
    debug_lines.append("")
    debug_lines.append(
        f"FINAL VLANS : {', '.join(sorted(vlans))}"
        if vlans
        else "NO VLAN FOUND"
    )

    set_vlan_debug("\n".join(debug_lines))
    return customer, sorted(vlans)
    #return customer, sorted(vlans,key=key)


def vlan_display_text(customer, vlans):
    return ((customer + " // ") if customer and vlans else "") + ("   ".join("#vlan " + v for v in vlans) if vlans else "No VLAN found")

# ============================================================
# EXCEL STYLING FUNCTION
# ============================================================

def style_excel_file(output_filename):
    """Apply SOC Dark-theme formatting to all sheets in the exported Excel file."""
    workbook = openpyxl.load_workbook(output_filename)

    # Dark Border Styling
    thin_border = Border(
        left=Side(style="thin", color="334155"),
        right=Side(style="thin", color="334155"),
        top=Side(style="thin", color="334155"),
        bottom=Side(style="thin", color="334155")
    )

    # Header Styling (Navy Blue background with White Font)
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")

    # Data Row Styling
    row_font = Font(name="Calibri", size=11, color="1E293B")
    row_alignment = Alignment(horizontal="center", vertical="center")

    for sheetname in workbook.sheetnames:
        sheet = workbook[sheetname]

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
        sheet.row_dimensions[1].height = 25

    workbook.save(output_filename)


# ============================================================
# FILE IO AND PROCESSING LOGIC
# ============================================================

def load_input_file(filename):
    extension = os.path.splitext(filename)[1].lower()
    if extension == ".csv":
        return pd.read_csv(filename, low_memory=False)
    elif extension == ".xlsx":
        return pd.read_excel(filename, engine="openpyxl")
    else:
        raise ValueError("Unsupported file type. Only CSV and XLSX are supported.")


def update_progress(percent):
    root.after(
        0,
        lambda: (
            progress_value.set(percent),
            progress_label.config(text=f"{int(percent)}%")
        )
    )


def process_log_file(filename, ticket_id, segregate_col):
    total_start = time.time()

    set_stage("Loading input log file...")
    df = load_input_file(filename)
    update_progress(20)

    set_stage("Inserting Numeric Ticket ID...")
    # Convert ticket ID to numeric integer value
    try:
        numeric_ticket_id = int(str(ticket_id).strip())
    except ValueError:
        numeric_ticket_id = pd.to_numeric(ticket_id, errors="coerce")

    if "Ticket ID" in df.columns:
        df.drop(columns=["Ticket ID"], inplace=True)
    df.insert(0, "Ticket ID", numeric_ticket_id)
    update_progress(35)

    set_stage("Normalizing blank values to 'null'...")
    # Replace NaN, whitespace-only, and empty strings with "null"
    df = df.fillna("null")
    df.replace(r"^\s*$", "null", regex=True, inplace=True)
    update_progress(50)

    # Determine structured target output filename
    output_filename = generate_formatted_filename(filename, ticket_id, segregate_col)

    set_stage("Generating XLSX payload...")
    with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
        if segregate_col and segregate_col in df.columns:
            set_stage("Exporting Primary Sheet ('All_Events')...")
            # 1. Add complete Process Log data as the primary first sheet
            df.to_excel(writer, sheet_name="All_Events", index=False)

            set_stage(f"Segregating sheets by '{segregate_col}'...")
            # 2. Add individual segregated sheets for each unique value
            grouped_df = df.copy()
            grouped_df[segregate_col] = grouped_df[segregate_col].astype(str)

            groups = grouped_df.groupby(segregate_col)
            sheet_names_used = {"All_Events"}

            for group_value, group_data in groups:
                # Sanitize sheet name for Excel guidelines
                safe_sheet_name = str(group_value)
                for char in [":", "\\", "/", "?", "*", "[", "]"]:
                    safe_sheet_name = safe_sheet_name.replace(char, "_")
                safe_sheet_name = safe_sheet_name[:30] if len(safe_sheet_name) > 30 else safe_sheet_name

                # Ensure unique sheet naming
                counter = 1
                unique_sheet_name = safe_sheet_name
                while unique_sheet_name in sheet_names_used:
                    unique_sheet_name = f"{safe_sheet_name[:25]}_{counter}"
                    counter += 1

                sheet_names_used.add(unique_sheet_name)
                group_data.to_excel(writer, sheet_name=unique_sheet_name, index=False)
        else:
            df.to_excel(writer, sheet_name="Processed_Events", index=False)

    update_progress(80)

    set_stage("Applying SOC terminal formatting...")
    style_excel_file(output_filename)
    update_progress(100)

    total_seconds = time.time() - total_start

    return {
        "output_filename": output_filename,
        "rows": len(df),
        "columns": len(df.columns),
        "elapsed": total_seconds,
        "customer": detect_vlans(df, filename)[0],
        "vlans": detect_vlans(df, filename)[1]
    }


# ============================================================
# GUI HANDLERS
# ============================================================

def select_file():
    global selected_filename, current_df, current_customer, current_vlans

    if processing_active:
        messagebox.showwarning("Process Running", "A log normalization process is active.")
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

    try:
        selected_filename = filename
        selected_file_label.config(text=filename, fg=APP_TEXT_MAIN)
        set_status("ANALYZING LOG FILE METADATA...", APP_ACCENT)

        current_df = load_input_file(filename)
        print("RUNNING VLAN DETECTION...")
        current_customer, current_vlans = detect_vlans(
            current_df,
            filename
        )

        print("CUSTOMER:", current_customer)
        print("VLANS:", current_vlans)

        display_text = vlan_display_text(
            current_customer,
            current_vlans
        )

        print("DISPLAY TEXT =", display_text)

        vlan_info_label.config(
            text=display_text,
            fg=APP_ACCENT if current_vlans else APP_WARNING
        )


        # DEBUG
        print("\n" + "=" * 80)
        print("FILE LOADED")
        print("FILENAME:", os.path.basename(filename))
        print("ROWS:", len(current_df))
        print("COLUMNS:", list(current_df.columns))
        print("\nFIRST 5 ROWS:")
        print(current_df.head())
        print("=" * 80)

        set_status("FILE LOADED // READY FOR PROCESSING", APP_SUCCESS)
        dropdown_selected(column_dropdown.get())

    except Exception as err:
        messagebox.showerror("File Error", f"Unable to read log file:\n{err}")
        set_status("FILE LOADING FAILED", APP_ERROR)


def dropdown_selected(selected_col):
    if current_df is None:
        unique_count_label.config(
            text="STATUS: Load a file to analyze unique values",
            fg=APP_TEXT_MUTED
        )
        return

    if selected_col in current_df.columns:
        valid_series = current_df[selected_col].dropna().astype(str).str.strip()
        valid_series = valid_series[~valid_series.str.lower().isin(["", "null", "nan", "none"])]

        unique_count = valid_series.nunique()

        if unique_count > 0:
            unique_count_label.config(
                text=f"UNIQUE VALUES IN '{selected_col}': {unique_count}",
                fg=APP_ACCENT
            )
        else:
            unique_count_label.config(
                text=f"UNIQUE VALUES IN '{selected_col}': no unique value",
                fg=APP_WARNING
            )
    else:
        unique_count_label.config(
            text=f"FIELD '{selected_col}' NOT PRESENT IN SOURCE LOG",
            fg=APP_ERROR
        )


def start_processing(segregate_multi_sheet=False):
    global processing_active, processing_started_at

    ticket_id = ticket_entry.get().strip()
    if not ticket_id:
        messagebox.showwarning("Missing Ticket ID", "Please enter a valid Ticket ID before proceeding.")
        return

    if not ticket_id.isdigit():
        messagebox.showwarning("Invalid Ticket ID", "Ticket ID must contain numbers only.")
        return

    if not selected_filename:
        messagebox.showwarning("Missing File", "Please select a target log file first.")
        return

    segregate_col = column_dropdown.get() if segregate_multi_sheet else None

    processing_active = True
    processing_started_at = time.time()

    select_button.config(state=tk.DISABLED, bg=APP_CARD_BG)
    execute_button.config(state=tk.DISABLED, bg=APP_CARD_BG)
    segregate_button.config(state=tk.DISABLED, bg=APP_CARD_BG)
    open_file_button.config(state=tk.DISABLED, bg=APP_CARD_BG)

    progress_value.set(0)
    progress_label.config(text="0%")
    set_status("INITIALIZING LOG NORMALIZATION...", APP_ACCENT)

    worker = threading.Thread(
        target=processing_worker,
        args=(selected_filename, ticket_id, segregate_col),
        daemon=True
    )
    worker.start()


def processing_worker(filename, ticket_id, segregate_col):
    try:
        result = process_log_file(filename, ticket_id, segregate_col)
        root.after(0, processing_completed, result)
    except PermissionError:
        error_message = "File Write Access Denied. Close the open spreadsheet and retry."
        root.after(0, processing_failed, error_message)
    except MemoryError:
        error_message = "System Out of Memory. Resource exhaustion prevents processing."
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
    execute_button.config(state=tk.NORMAL, bg=APP_BTN_PRIMARY, fg="white")
    segregate_button.config(state=tk.NORMAL, bg=APP_BTN_PRIMARY, fg="white")
    open_file_button.config(state=tk.NORMAL, bg=APP_BTN_PRIMARY, fg="white")

    set_status("EXECUTION SUCCESS", APP_SUCCESS)

    messagebox.showinfo(
        "Operation Complete",
        "Log Normalization Finished Successfully.\n\n"
        f"Rows Processed: {result['rows']}\n"
        f"Total Fields: {result['columns']}\n"
        f"Execution Time: {result['elapsed']:.2f}s\n\n"
        f"Output Target:\n{result['output_filename']}"
    )


def processing_failed(error_message):
    global processing_active

    processing_active = False
    select_button.config(state=tk.NORMAL, bg=APP_BTN_PRIMARY, fg="white")
    execute_button.config(state=tk.NORMAL, bg=APP_BTN_PRIMARY, fg="white")
    segregate_button.config(state=tk.NORMAL, bg=APP_BTN_PRIMARY, fg="white")

    set_status("EXECUTION ABORTED // ERROR ENCOUNTERED", APP_ERROR)
    messagebox.showerror("Normalization Failure", error_message)


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
root.geometry("700x625")
root.resizable(False, False)
root.configure(bg=APP_BACKGROUND)
progress_value = tk.DoubleVar(value=0)
root.protocol("WM_DELETE_WINDOW", exit_application)

# ttk Style Customization for Progressbar and Combobox
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
style.configure(
    "TCombobox",
    fieldbackground="#0F172A",
    background=APP_CARD_BG,
    foreground=APP_TEXT_MAIN,
    arrowcolor=APP_ACCENT
)

main_frame = tk.Frame(root, bg=APP_BACKGROUND)
main_frame.pack(fill=tk.BOTH, expand=True, padx=24, pady=20)

# Header Banner
title_label = tk.Label(
    main_frame,
    text="🛡️ CC_NORM // LOG NORMALIZER",
    bg=APP_CARD_BG,
    fg=APP_ACCENT,
    font=("Consolas", 15, "bold"),
    pady=10,
    relief=tk.FLAT
)
title_label.pack(fill=tk.X)

# Subtitle Info
instruction_label = tk.Label(
    main_frame,
    text="Parses raw logs, injects Ticket ID column, evaluates unique value distributions, and segregates data across multiple sheets.",
    bg=APP_BACKGROUND,
    fg=APP_TEXT_MUTED,
    font=("Segoe UI", 9),
    wraplength=550,
    justify=tk.CENTER
)
instruction_label.pack(pady=(10, 10))

# Selection & Settings Container Card
card_frame = tk.Frame(main_frame, bg=APP_CARD_BG, padx=15, pady=12)
card_frame.pack(fill=tk.X, pady=(0, 10))

# Ticket ID Input Row
ticket_label = tk.Label(
    card_frame,
    text="ENTER TICKET ID:",
    bg=APP_CARD_BG,
    fg=APP_TEXT_MUTED,
    font=("Consolas", 9, "bold")
)
ticket_label.pack(anchor="w", pady=(0, 2))

ticket_entry = tk.Entry(
    card_frame,
    bg="#0F172A",
    fg=APP_TEXT_MAIN,
    insertbackground=APP_ACCENT,
    font=("Consolas", 10),
    relief=tk.FLAT,
    bd=5
)
ticket_entry.pack(fill=tk.X, pady=(0, 10))

# Target Log Display
selected_file_title = tk.Label(
    card_frame,
    text="TARGET LOG SOURCE:",
    bg=APP_CARD_BG,
    fg=APP_TEXT_MUTED,
    font=("Consolas", 9, "bold")
)
selected_file_title.pack(anchor="w", pady=(0, 2))

selected_file_label = tk.Label(
    card_frame,
    text="No log file loaded...",
    bg="#0F172A",
    fg="#475569",
    font=("Consolas", 9),
    relief=tk.FLAT,
    anchor="w",
    padx=10,
    pady=6,
    wraplength=520,
    justify=tk.LEFT
)
selected_file_label.pack(fill=tk.X, pady=(0, 10))

# Dropdown Analysis Section
dropdown_title = tk.Label(
    card_frame,
    text="SELECT FIELD FOR UNIQUE VALUE ANALYSIS:",
    bg=APP_CARD_BG,
    fg=APP_TEXT_MUTED,
    font=("Consolas", 9, "bold")
)
dropdown_title.pack(anchor="w", pady=(0, 2))

column_dropdown = ttk.Combobox(
    card_frame,
    values=ANALYSIS_COLUMNS,
    state="readonly",
    font=("Segoe UI", 9)
)
column_dropdown.set(ANALYSIS_COLUMNS[0])
column_dropdown.pack(fill=tk.X, pady=(0, 5))
column_dropdown.bind("<<ComboboxSelected>>", lambda e: dropdown_selected(column_dropdown.get()))

unique_count_label = tk.Label(
    card_frame,
    text="STATUS: Load a file to analyze unique values",
    bg=APP_CARD_BG,
    fg=APP_TEXT_MUTED,
    font=("Consolas", 9, "bold")
)
unique_count_label.pack(anchor="w", pady=(2, 0))

# Primary Control Buttons
button_frame = tk.Frame(main_frame, bg=APP_BACKGROUND)
button_frame.pack(pady=(5, 10))

select_button = tk.Button(
    button_frame,
    text="Select Log File",
    command=select_file,
    bg=APP_BTN_PRIMARY,
    fg="white",
    activebackground=APP_BTN_HOVER,
    activeforeground="white",
    font=("Segoe UI", 9, "bold"),
    width=14,
    relief=tk.FLAT,
    cursor="hand2"
)
select_button.grid(row=0, column=0, padx=4, pady=2)

execute_button = tk.Button(
    button_frame,
    text="Process Log",
    command=lambda: start_processing(segregate_multi_sheet=False),
    bg=APP_BTN_PRIMARY,
    fg="white",
    activebackground=APP_BTN_HOVER,
    activeforeground="white",
    font=("Segoe UI", 9, "bold"),
    width=14,
    relief=tk.FLAT,
    cursor="hand2"
)
execute_button.grid(row=0, column=1, padx=4, pady=2)

segregate_button = tk.Button(
    button_frame,
    text="Make Multiple Sheet",
    command=lambda: start_processing(segregate_multi_sheet=True),
    bg=APP_BTN_PRIMARY,
    fg="white",
    activebackground=APP_BTN_HOVER,
    activeforeground="white",
    font=("Segoe UI", 9, "bold"),
    width=16,
    relief=tk.FLAT,
    cursor="hand2"
)
segregate_button.grid(row=0, column=2, padx=4, pady=2)

open_file_button = tk.Button(
    button_frame,
    text="Open Export File",
    command=open_output_file,
    bg=APP_CARD_BG,
    fg=APP_TEXT_MUTED,
    activebackground=APP_BTN_HOVER,
    activeforeground="white",
    font=("Segoe UI", 9, "bold"),
    width=14,
    state=tk.DISABLED,
    relief=tk.FLAT,
    cursor="hand2"
)
open_file_button.grid(row=0, column=3, padx=4, pady=2)
exit_button = tk.Button(button_frame, text="Exit", command=exit_application, bg=APP_ERROR, fg="white", activebackground="#B91C1C", activeforeground="white", font=("Segoe UI", 9, "bold"), width=10, relief=tk.FLAT, cursor="hand2")
exit_button.grid(row=0, column=4, padx=4, pady=2)

# Progress Diagnostics
progress_bar = ttk.Progressbar(
    main_frame,
    variable=progress_value,
    maximum=100,
    mode="determinate",
    style="Custom.Horizontal.TProgressbar"
)
progress_bar.pack(fill=tk.X, pady=(5, 2))

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
status_label.pack(pady=(2, 2))
vlan_info_label = tk.Label(main_frame, text="No VLAN found", bg=APP_BACKGROUND, fg=APP_WARNING, font=("Consolas", 10, "bold"), wraplength=640, justify=tk.CENTER)
vlan_info_label.pack(pady=(0, 5))

vlan_debug_label = tk.Label(
    main_frame,
    text="",
    bg=APP_BACKGROUND,
    fg=APP_TEXT_MUTED,
    font=("Consolas", 8),
    justify=tk.LEFT,
    anchor="w",
    wraplength=650
)
vlan_debug_label.pack(fill=tk.X, pady=(0, 5))

developer_label = tk.Label(
    main_frame,
    text="DEV // Gourab Sarkar",
    bg=APP_BACKGROUND,
    fg="#475569",
    font=("Consolas", 8)
)
developer_label.pack(anchor="e")


if __name__ == "__main__":
    root.mainloop()