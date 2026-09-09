# Normalizer 2.0

Normalizer is a multi-threaded Python desktop GUI application from [https://github.com/zohan205/Normalizer-2.0.git](https://github.com/zohan205/Normalizer-2.0.git) designed to clean, process, and standardize security log exports.

## Features

- **Strict Extension Filtering**: Accepts `.csv` and `.xlsx` formats.
- **Universal CSV Engine**: Converts flat `.csv` logs for uniform processing.
- **Dynamic Structural Layout Shifts**: Positions `rawEvent` in Column A and `eventTime` in Column B.
- **Smart Data Clean-Up**: Drops unwanted tracking fields and empty/null entries, while auto-fitting column widths safely.

## Requirements

- **Python 3.x** and packages including `openpyxl`, `tkinter`, and standard libraries.

## Installation & Usage

1. Clone and navigate:
    ```bash
    git clone https://github.com/zohan205/Normalizer-2.0.git
    cd Normalizer-2.0
    ```
2. Run the application:
    ```bash
    python Normalizer_2.py
    ```
3. Compile to executable:
    ```bash
    pyinstaller --onefile --noconsole --hidden-import=tkinter.font Normalizer_2.py
    ```
