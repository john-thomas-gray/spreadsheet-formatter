# Spreadsheet File Converter

This project provides tools for converting spreadsheet files between XLS and CSV formats, specifically tailored for processing dialog lists with particular focus on foreign language
subtitle text.

## Setup Instructions

Before using the conversion scripts, set up and activate a Python virtual environment and install the necessary dependencies.

### 1. Activate Virtual Environment

    **macOS and Linux**

    ```bash
    python3 -m venv env
    source env/bin/activate
    ```
    **Windows:**

    ```bash
    python -m venv env
    .\env\Scripts\activate
    ```

### 2. Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```
## Use Cases:

### process_xls.py

Convert xls dialog list to a properly formatted csv.
    - Removes English dialog, leaving only Spanish dialog
    - Adds character names
    - Removes dialog introductions (i.e. "Man 1 says in Spanish")
    - Returns csv document containing only relevant rows

Instructions:
    1. Put xls files into xls_in directory
    2. Change current directory to project's root directory
    3. In terminal, enter:
        python process_xls.py
    4. Retrieve formatted files from csv_out

### process_csv_to_xls.py

Convert csv dialog list to a properly formatted xlsx.
    - Formats file to client's specifications
    - Returns xlsx document containing only relevant columns

Instructions:
    1. Put csv files into csv_in directory
    2. Change current directory to project's root directory
    3. In terminal, enter:
        python process_csv_to_xls.py
    4. Retrieve formatted files from xls_out
