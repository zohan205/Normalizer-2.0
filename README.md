# Normalizer 2.0

Normalizer is a multi-threaded Python desktop application designed to clean, process, and standardize security log exports. It handles empty datasets, manages metadata shifts dynamically, strips useless tracking fields, and structures data with a user-friendly Graphical User Interface (GUI).

## Features

- **Strict Extension Filtering**: Protects processing pipelines by accepting only `.csv` and `.xlsx` formats. 
- **Universal CSV Engine**: Automatically imports flat `.csv` logs into an in-memory worksheet format to apply styles and filters identically to Excel workbooks.
- **Dynamic Structural Layout Shifts**: Automates structural column prioritization by shifting the critical `rawEvent` data blob directly into the 1st column (Column A) and the `eventTime` clock metrics directly into the 2nd column (Column B).
- **Metadata Title Reconstruction**: Parses dynamic reference markers (such as mapping header strings ending in "Name") to clean up system variables safely.
- **Smart Data Clean-Up**: Identifies and drops predefined structural clutter (`rawEventHash`, `customerURI`, etc.) alongside any metrics loaded entirely with empty fields or string `"null"` text values.
- **Capped Auto-Fit Width Adjustments**: Intelligently auto-sizes regular columns while setting a custom formatting safety ceiling for massive json blobs like `rawEvent` to prevent broken workbook layouts.

## Requirements

- **Python 3.x**
- **Required Python Packages**:
  - `openpyxl`: For in-memory calculation and layout formatting.
  - `tkinter`: Embedded window environment module.
  - `csv`, `datetime`, `sys`, `os`, `time`, `threading`: Local environment modules.

## Installation

1. **Clone the Repository**:
    ```bash
    git clone https://github.com/zohan205/Normalizer.git
    ```

2. **Navigate to the Project Directory**:
    ```bash
    cd Normalizer
    ```

3. **Create and Activate a Virtual Environment**:
    - Windows:
        ```bash
        python -m venv venv
        .\venv\Scripts\activate
        ```
    - macOS/Linux:
        ```bash
        python3 -m venv venv
        source venv/bin/activate
        ```

4. **Install Dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

### Running the Application

You can execute the processing module as raw Python script text tracks directly from your active shell environment terminal:
```bash
python Normalizer_2.py
```

### Compiling to a Standalone Executable (.exe)

To bundle the Python dependencies, internal file pickers, and asset structures into a single desktop execution file that runs on systems without a Python runtime environment installed, execute the following command from your active directory:

```bash
pyinstaller --onefile --noconsole --hidden-import=tkinter.font test110.py
```

*Note: The `--hidden-import=tkinter.font` flag is mandatory to force the compiler tracks to anchor your application graphics layouts cleanly and prevent immediate background operating system launch crashes.*
