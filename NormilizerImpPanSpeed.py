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
    """
    Listen for user input.

    Enter x and press Enter to terminate the program.
    """
    while not exit_requested.is_set():
        try:
            line = sys.stdin.readline()

            if line and line.strip().lower() == "x":
                print("[INFO] 'x' entered by user. Terminating program...")
                os._exit(0)

        except Exception:
            break


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

    Excel column headers cannot be empty. This function also replaces
    line breaks and tab characters with spaces.
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
    """
    Return a safe output column name.

    Normally, the metadata Name value is used directly.

    Examples:
    - Event Count
    - TEST
    - ICMP Code

    If an unrelated original column already uses the same name,
    a numbered suffix is added to prevent overwriting the original data.
    """
    requested_name = clean_filename_column_name(requested_name)

    if not requested_name:
        return None

    # Reuse an output column that was already generated.
    if requested_name in generated_columns:
        return requested_name

    # Use the requested name if it does not exist in the original data.
    if requested_name not in original_columns:
        return requested_name

    # The original source value column will be removed later.
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

    Example input:

        aisaacNum1Name    aisaacNum1
        Event Count      84831
        Event Count      84840
        TEST             84795
        TEST             84804

    Example output:

        Event Count      TEST
        84831
        84840
                         84795
                         84804

    Important:
    - No rows are added.
    - No rows are deleted.
    - Row order is not changed.
    - Non-applicable cells remain empty.
    - New columns are added once using pd.concat().
    - This avoids DataFrame fragmentation.
    """
    original_index = df.index
    original_row_count = len(df)
    original_columns = list(df.columns)
    original_column_set = set(original_columns)

    metadata_pairs = []

    # Identify all Name/value column pairs.
    for name_column in original_columns:
        name_column_text = str(name_column)

        if not name_column_text.endswith("Name"):
            continue

        # Examples:
        # aisaacNum1Name -> aisaacNum1
        # aisaacStr1Name -> aisaacStr1
        # colStr1Name    -> colStr1
        value_column = name_column_text[:-4]

        if value_column in df.columns:
            metadata_pairs.append(
                (name_column, value_column)
            )

    if not metadata_pairs:
        return df

    # Store generated output Series here.
    # No columns are inserted into df inside the loop.
    generated_columns = {}

    # Store columns that will be removed after expansion.
    columns_to_drop = []

    for name_column, value_column in metadata_pairs:
        cleaned_names, valid_name_mask = clean_name_series(
            df[name_column]
        )

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

            matching_rows = (
                valid_name_mask
                & cleaned_names.eq(unique_name)
            )

            # Create the output Series only in memory.
            # This does not insert anything into the original DataFrame.
            if output_column not in generated_columns:
                generated_columns[output_column] = pd.Series(
                    pd.NA,
                    index=original_index,
                    dtype="object"
                )

            target_series = generated_columns[output_column]
            source_values = df[value_column]

            # Check if the destination is empty.
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

            # Copy the corresponding value into the matching rows.
            target_series.loc[update_mask] = source_values.loc[
                update_mask
            ].values

            generated_columns[output_column] = target_series

        columns_to_drop.extend(
            [name_column, value_column]
        )

    # Remove duplicates while preserving order.
    columns_to_drop = list(dict.fromkeys(columns_to_drop))

    existing_columns_to_drop = [
        column
        for column in columns_to_drop
        if column in df.columns
    ]

    # Drop original metadata Name/value columns.
    base_df = df.drop(
        columns=existing_columns_to_drop
    )

    # Convert all generated columns into one DataFrame.
    if generated_columns:
        generated_df = pd.DataFrame(
            generated_columns,
            index=original_index
        )

        # Add every generated column in one operation.
        result_df = pd.concat(
            [base_df, generated_df],
            axis=1
        )
        result_df = result_df.copy()
        

    else:
        result_df = base_df

    # Create a clean contiguous DataFrame.
    result_df = result_df.copy()

    # Safety checks.
    if len(result_df) != original_row_count:
        raise RuntimeError(
            "Metadata expansion changed the number of rows."
        )

    if not result_df.index.equals(original_index):
        raise RuntimeError(
            "Metadata expansion changed the DataFrame index."
        )

    return result_df


def drop_empty_columns(df):
    """
    Remove columns that contain only:
    - Empty values
    - null
    - nan
    - none
    """
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
        df = df.drop(
            columns=columns_to_drop
        )

    return df


def rearrange_columns(df):
    """
    Put rawEvent first and eventTime second.

    All other columns keep their existing order.
    """
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
    """
    Load CSV or XLSX input.
    """
    extension = os.path.splitext(filename)[1].lower()

    if extension == ".csv":
        return pd.read_csv(
            filename,
            low_memory=False
        )

    if extension == ".xlsx":
        return pd.read_excel(
            filename,
            engine="openpyxl"
        )

    raise ValueError(
        "Unsupported file type. Only CSV and XLSX are supported."
    )


def save_dataframe_to_excel(df, output_filename):
    """
    Save the cleaned DataFrame to Excel.
    """
    df.to_excel(
        output_filename,
        index=False,
        engine="openpyxl"
    )


def style_excel_file(output_filename):
    """
    Apply Excel formatting.
    """
    workbook = openpyxl.load_workbook(output_filename)
    sheet = workbook.active
    sheet.title = "Tiggered Events"

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

    # Style header.
    for cell in sheet[1]:
        cell.border = thin_border
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment

    # Style all data cells.
    row_alignment = Alignment(
        horizontal='center',
        vertical='center'
    )

    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.border = thin_border
            cell.font = row_font
            cell.alignment = row_alignment

    # Adjust column widths.
    for column_cells in sheet.columns:
        max_length = 0

        column_letter = openpyxl.utils.get_column_letter(
            column_cells[0].column
        )

        column_header = column_cells[0].value

        for cell in column_cells:
            if cell.value is not None:
                value_length = len(str(cell.value))

                if value_length > max_length:
                    max_length = value_length

        adjusted_width = (max_length + 2) * 1.2

        if column_header == "rawEvent":
            sheet.column_dimensions[column_letter].width = 50
        else:
            sheet.column_dimensions[column_letter].width = max(
                min(adjusted_width, 40),
                10
            )

    # Keep the header visible.
    sheet.freeze_panes = "A2"

    # Enable filtering.
    if sheet.max_row >= 1 and sheet.max_column >= 1:
        sheet.auto_filter.ref = None

    # Set header row height.
    sheet.row_dimensions[1].height = 22
    workbook.save(output_filename)


def process_log_file(filename):
    """
    Main processing function.
    """
    total_start = time.time()

    try:
        # --- 1. Load File ---
        load_start = time.time()

        df = load_input_file(filename)

        print(
            f"\n[INFO] File loaded in "
            f"{time.time() - load_start:.2f} seconds."
        )

        original_row_count = len(df)
        original_column_count = len(df.columns)

        print(f"[INFO] Input rows: {original_row_count}")
        print(f"[INFO] Input columns: {original_column_count}")

        # --- 2. Rearrange Columns ---
        df = rearrange_columns(df)

        # --- 3. Drop Predefined Unwanted Columns ---
        unwanted_columns = [
            "rawEventHash",
            "aisaacReceivedTime",
            "customerURI",
            "cenNifiReceiptTime",
            "logFilterKafkaInTime",
            "logFilterInTime"
        ]

        existing_unwanted_columns = [
            column
            for column in unwanted_columns
            if column in df.columns
        ]

        if existing_unwanted_columns:
            df = df.drop(
                columns=existing_unwanted_columns
            )

        # --- 4. Expand Metadata Columns ---
        metadata_start = time.time()

        df = expand_metadata_columns_fast(df)

        print(
            f"[INFO] Metadata columns expanded in "
            f"{time.time() - metadata_start:.2f} seconds."
        )

        # Confirm row count did not change.
        if len(df) != original_row_count:
            raise RuntimeError(
                "Row count changed after metadata expansion."
            )

        # --- 5. Drop Empty Columns ---
        null_scan_start = time.time()

        df = drop_empty_columns(df)

        print(
            f"[INFO] Empty columns removed in "
            f"{time.time() - null_scan_start:.2f} seconds."
        )

        # Final row check.
        if len(df) != original_row_count:
            raise RuntimeError(
                "Row count changed during processing."
            )

        # --------------------------------------
        # Convert all blank values to "null"
        # --------------------------------------

        df = df.fillna("null")

        df.replace(
            to_replace=r'^\s*$',
            value='null',
            regex=True,
            inplace=True
        )

        # --- 6. Create Output Filename ---
        timestamp = datetime.now().strftime(
            "%Y-%m-%d_%H-%M-%S"
        )

        base_path, _ = os.path.splitext(filename)

        output_filename = (
            f"{base_path}_{timestamp}.xlsx"
        )

        # --- 7. Save DataFrame ---
        write_start = time.time()

        save_dataframe_to_excel(
            df,
            output_filename
        )

        print(
            f"[INFO] Excel data written in "
            f"{time.time() - write_start:.2f} seconds."
        )

        # --- 8. Apply Excel Styling ---
        style_start = time.time()

        style_excel_file(output_filename)

        print(
            f"[INFO] Excel styling completed in "
            f"{time.time() - style_start:.2f} seconds."
        )

        print("\n[SUCCESS] Processing completed.")
        print(f"Rows preserved: {len(df)}")
        print(f"Final columns: {len(df.columns)}")
        print(
            f"Total elapsed time: "
            f"{time.time() - total_start:.2f} seconds"
        )
        print(f"Saved as: {output_filename}")

    except PermissionError:
        print(
            "\n[ERROR] Permission denied while writing the output file."
        )
        print(
            "Close the output Excel file if it is already open, "
            "then run the script again."
        )

    except MemoryError:
        print(
            "\n[ERROR] The computer does not have enough available "
            "memory to process this file."
        )

    except Exception as error:
        print(f"\n[ERROR] Processing failed: {error}")


if __name__ == "__main__":
    print("=== Log Normalizer (developed by Gourab) ===")

    if len(sys.argv) > 1:
        filename = sys.argv[1]
    else:
        filename = input(
            "Enter CSV or XLSX file path: "
        )

    filename = (
        filename
        .strip()
        .strip('"')
        .strip("'")
    )

    if not filename:
        print("Error: No file path was provided.")
        sys.exit(1)

    if not os.path.exists(filename):
        print(f"Error: File '{filename}' was not found.")
        sys.exit(1)

    if not os.path.isfile(filename):
        print(f"Error: '{filename}' is not a file.")
        sys.exit(1)

    supported_extensions = [".csv", ".xlsx"]
    file_extension = os.path.splitext(filename)[1].lower()

    if file_extension not in supported_extensions:
        print(
            "Error: Unsupported file type. "
            "Only CSV and XLSX files are supported."
        )
        sys.exit(1)

    # Start cancellation listener.
    threading.Thread(
        target=input_listener,
        daemon=True
    ).start()

    # Start processing worker.
    worker = threading.Thread(
        target=process_log_file,
        args=(filename,)
    )

    worker.start()

    start_wait = time.time()
    warning_displayed = False

    while worker.is_alive():
        worker.join(timeout=1.0)

        if (
            time.time() - start_wait > 60
            and not warning_displayed
        ):
            print(
                "\n[NOTICE] The program has been running "
                "for over 60 seconds."
                
            )
            print(
                "To terminate the program, enter 'x' "
                "and press Enter: \n",
                end="",
                flush=True
            )

            warning_displayed = True

    exit_requested.set()