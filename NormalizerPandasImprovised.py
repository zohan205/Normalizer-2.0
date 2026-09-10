import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import sys
import os
import time
import threading
import warnings
from pandas.errors import PerformanceWarning

warnings.simplefilter("ignore", PerformanceWarning)


exit_requested = threading.Event()


def input_listener():
    """Listen for 'x' input and terminate the program when requested."""
    while not exit_requested.is_set():
        try:
            line = sys.stdin.readline()

            if line and line.strip().lower() == "x":
                print("\n[INFO] 'x' entered by user. Terminating program...")
                os._exit(0)

        except Exception:
            break


def clean_metadata_name_series(series):
    """
    Clean metadata Name values.

    Returns:
        cleaned_names: Trimmed string values
        valid_mask: True only for usable Name values
    """
    cleaned_names = series.astype("string").str.strip()

    valid_mask = (
        series.notna()
        & cleaned_names.notna()
        & cleaned_names.ne("")
        & cleaned_names.str.lower().ne("null")
        & cleaned_names.str.lower().ne("nan")
        & cleaned_names.str.lower().ne("none")
    )

    return cleaned_names, valid_mask


def get_unique_column_name(existing_columns, requested_name, source_value_column):
    """
    Return a safe unique output column name.

    Normally the output column uses the metadata Name value directly,
    such as 'Event Count' or 'TEST'.

    If that name already belongs to an unrelated original column,
    create a unique version to prevent overwriting existing data.
    """
    requested_name = str(requested_name).strip()

    if requested_name not in existing_columns:
        return requested_name

    # It is safe to reuse the source value column's own name because
    # that source column will be removed after expansion.
    if requested_name == source_value_column:
        return requested_name

    counter = 2
    new_name = f"{requested_name}_{counter}"

    while new_name in existing_columns:
        counter += 1
        new_name = f"{requested_name}_{counter}"

    return new_name


def expand_metadata_columns(df):
    """
    Convert each metadata Name/value pair into separate output columns.

    Example input:
        aisaacNum1Name    aisaacNum1
        Event Count      84831
        TEST             84795

    Example output:
        Event Count      TEST
        84831            empty
        empty            84795

    The DataFrame index and row positions are not changed.
    """
    columns_to_drop = []

    # Use a fixed copy because new columns are added during processing.
    original_columns = list(df.columns)

    for name_column in original_columns:
        name_column_text = str(name_column)

        if not name_column_text.endswith("Name"):
            continue

        # Examples:
        # aisaacNum1Name -> aisaacNum1
        # aisaacStr1Name -> aisaacStr1
        # colStr1Name    -> colStr1
        value_column = name_column_text[:-4]

        if value_column not in df.columns:
            continue

        cleaned_names, valid_name_mask = clean_metadata_name_series(
            df[name_column]
        )

        unique_names = (
            cleaned_names.loc[valid_name_mask]
            .drop_duplicates()
            .tolist()
        )

        # Keep the old behavior for a single unique Name.
        # Also support multiple unique Name values.
        created_columns = {}

        for unique_name in unique_names:
            unique_name = str(unique_name).strip()

            if not unique_name:
                continue

            matching_rows = (
                valid_name_mask
                & cleaned_names.eq(unique_name)
            )

            # Reuse the same output column if this Name was already
            # created by another Name/value pair.
            if unique_name in created_columns:
                output_column = created_columns[unique_name]

            elif (
                unique_name in df.columns
                and unique_name not in original_columns
            ):
                output_column = unique_name
                created_columns[unique_name] = output_column

            else:
                output_column = get_unique_column_name(
                    set(df.columns),
                    unique_name,
                    value_column
                )
                created_columns[unique_name] = output_column

            if output_column not in df.columns:
                df[output_column] = pd.NA

            # Only fill matching rows.
            # Other rows remain empty and row numbers are preserved.
            source_values = df.loc[matching_rows, value_column]

            # Do not overwrite existing values if the same output column
            # has already been populated from another metadata pair.
            destination_is_empty = (
                df.loc[matching_rows, output_column].isna()
                | df.loc[matching_rows, output_column]
                    .astype("string")
                    .str.strip()
                    .isin(["", "null", "nan", "none"])
            )

            rows_to_update = destination_is_empty[
                destination_is_empty
            ].index

            df.loc[rows_to_update, output_column] = df.loc[
                rows_to_update,
                value_column
            ].values

        # Remove the original metadata Name/value pair after expansion.
        columns_to_drop.extend([name_column, value_column])

    # Remove duplicate column names from the drop list while
    # preserving the list order.
    columns_to_drop = list(dict.fromkeys(columns_to_drop))

    existing_drop_columns = [
        column
        for column in columns_to_drop
        if column in df.columns
    ]

    if existing_drop_columns:
        df.drop(
            columns=existing_drop_columns,
            inplace=True
        )

    # De-fragment dataframe
    df = df.copy()

    return df


def process_log_file(filename):
    total_start = time.time()
    extension = os.path.splitext(filename)[1].lower()

    # --- 1. Load File Using Pandas ---
    try:
        if extension == ".csv":
            df = pd.read_csv(
                filename,
                low_memory=False
            )

        elif extension == ".xlsx":
            df = pd.read_excel(
                filename,
                engine="openpyxl"
            )

        else:
            print("Error: Unsupported file type.")
            print("Supported file types: .csv and .xlsx")
            return

    except Exception as error:
        print(f"Error loading file: {error}")
        return

    original_row_count = len(df)

    # --- 2. Rearrange Columns ---
    # rawEvent must be first and eventTime must be second.
    columns = list(df.columns)

    if "rawEvent" in columns:
        columns.remove("rawEvent")

    if "eventTime" in columns:
        columns.remove("eventTime")

    front_columns = []

    if "rawEvent" in df.columns:
        front_columns.append("rawEvent")

    if "eventTime" in df.columns:
        front_columns.append("eventTime")

    df = df[front_columns + columns]

    # --- 3. Drop Predefined Unwanted Columns ---
    drop_column_names = [
        "rawEventHash",
        "aisaacReceivedTime",
        "customerURI",
        "cenNifiReceiptTime",
        "logFilterKafkaInTime",
        "logFilterInTime"
    ]

    df.drop(
        columns=[
            column
            for column in drop_column_names
            if column in df.columns
        ],
        inplace=True
    )

    # --- 4. Expand Metadata Name/Value Columns ---
    #
    # Example:
    # aisaacNum1Name = Event Count -> Event Count column
    # aisaacNum1Name = TEST        -> TEST column
    #
    # The corresponding aisaacNum1 value is copied into the
    # applicable column on the same row.
    df = expand_metadata_columns(df)

    # Ensure dataframe is contiguous in memory
    df = df.copy()

    # Safety check: metadata expansion must not change row count.
    if len(df) != original_row_count:
        print(
            "[ERROR] Row count changed unexpectedly during "
            "metadata expansion."
        )
        return

    # --- 5. Drop Empty or Purely Null Columns ---
    null_columns_to_drop = []

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
            null_columns_to_drop.append(column)

    if null_columns_to_drop:
        df.drop(
            columns=null_columns_to_drop,
            inplace=True
        )

    # --- 6. Save Cleaned Data to Excel ---
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    base_path, _ = os.path.splitext(filename)
    output_filename = f"{base_path}_{timestamp}.xlsx"

    try:
        df.to_excel(
            output_filename,
            index=False,
            engine="openpyxl"
        )

    except Exception as error:
        print(f"Error writing Excel file: {error}")
        return

    # --- 7. Apply Excel Styling ---
    try:
        workbook = openpyxl.load_workbook(output_filename)
        sheet = workbook.active

        thin_border = openpyxl.styles.Border(
            left=openpyxl.styles.Side(style="thin"),
            right=openpyxl.styles.Side(style="thin"),
            top=openpyxl.styles.Side(style="thin"),
            bottom=openpyxl.styles.Side(style="thin")
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

        # Style header row.
        for cell in sheet[1]:
            cell.border = thin_border
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment

        # Style data rows.
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.border = thin_border
                cell.font = row_font

        # Adjust column widths.
        for column_cells in sheet.columns:
            max_length = 0

            column_letter = openpyxl.utils.get_column_letter(
                column_cells[0].column
            )

            column_header = column_cells[0].value

            for cell in column_cells:
                if cell.value is not None:
                    cell_length = len(str(cell.value))
                    max_length = max(max_length, cell_length)

            adjusted_width = (max_length + 2) * 1.2

            if column_header == "rawEvent":
                sheet.column_dimensions[column_letter].width = 50
            else:
                sheet.column_dimensions[column_letter].width = max(
                    min(adjusted_width, 40),
                    10
                )

        # Keep the header visible while scrolling.
        sheet.freeze_panes = "A2"

        # Enable filters for all output columns.
        sheet.auto_filter.ref = sheet.dimensions

        workbook.save(output_filename)

    except Exception as error:
        print(f"Error styling Excel file: {error}")
        return

    print(
        f"\n[SUCCESS] Total Elapsed Time: "
        f"{time.time() - total_start:.2f} seconds"
    )
    print(f"Rows preserved: {original_row_count}")
    print(f"Saved as: {output_filename}")


if __name__ == "__main__":
    print("=== Log Normalizer (developed by Gourab) ===")

    if len(sys.argv) > 1:
        filename = sys.argv[1]
    else:
        filename = input("Enter file path: ")

    filename = filename.strip().strip('"').strip("'")

    if not filename or not os.path.exists(filename):
        print(f"Error: File '{filename}' not found.")
        sys.exit(1)

    # Start listener thread for the 'x' cancellation option.
    threading.Thread(
        target=input_listener,
        daemon=True
    ).start()

    # Run file processing in a worker thread.
    worker = threading.Thread(
        target=process_log_file,
        args=(filename,)
    )

    worker.start()

    start_wait = time.time()
    warned = False

    while worker.is_alive():
        worker.join(timeout=1.0)

        if time.time() - start_wait > 60 and not warned:
            print(
                "\n[NOTICE] Program has been running "
                "for over 60 seconds."
            )
            print(
                "To exit, enter 'x' and press Enter: ",
                end="",
                flush=True
            )
            warned = True

    exit_requested.set()