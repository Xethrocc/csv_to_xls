# CSV to XLS Converter

This script converts CSV (Comma Separated Values) files into XLS (Excel 97-2003) format.

It can operate in two modes:

1.  **Folder Mode:** Converts all `.csv` files found in a specified input directory (default: `input/`) and saves the corresponding `.xls` files to an output directory (default: `output/`).
2.  **Single File Mode:** Converts a specific input CSV file to a specified output XLS file.

The script automatically tries to detect whether the CSV delimiter is a comma (`,`) or a semicolon (`;`).

## Project Structure

```
.
├── src/                   # Contains the main conversion logic
│   └── converter.py
├── input/                 # Default directory for input CSV files
├── output/                # Default directory for output XLS files (created if needed)
├── venv/                  # Example Python virtual environment (ignored by Git)
├── .gitignore             # Specifies files/folders for Git to ignore
├── README.md              # This file
├── DE_README.md           # German version of the README
├── requirements.txt       # Lists dependencies (xlwt)
├── setup_and_run.bat      # Windows batch script for setup and execution
└── start_conversion.py    # Easy way to run the converter from the root
```

## Setup

1.  **Clone the repository (if applicable):**
    ```bash
    git clone <repository-url>
    cd <repository-folder>
    ```

2.  **Create a virtual environment (Recommended):**
    ```bash
    python -m venv venv
    ```
    *   On Windows: `venv\Scripts\activate`
    *   On macOS/Linux: `source venv/bin/activate`

3.  **Install dependencies:**
    The script requires the `xlwt` library. You should install it using the `requirements.txt` file.
    ```bash
    pip install -r requirements.txt
    ```
    *(Alternatively, install manually: `pip install xlwt`)*

4.  **Prepare Input Files:**
    *   Place your `.csv` files into the `input/` directory (or create it if it doesn't exist).

## Usage

You can run the conversion using the `start_conversion.py` script from the project's root directory, or use the `setup_and_run.bat` script on Windows which handles setup and execution.

**Using `setup_and_run.bat` (Windows):**

*   Double-click the `setup_and_run.bat` file.
*   It will attempt to create a virtual environment (if needed), install dependencies from `requirements.txt`, and then run the converter in Folder Mode.
*   Follow the prompts in the console window. Administrator privileges might be needed if Python installation via winget is attempted.

**Using `start_conversion.py` (Manual):**

Make sure you have activated your virtual environment and installed dependencies first.

**1. Folder Mode (Default):**

*   Ensure your CSV files are in the `input/` directory.
*   Run the start script:
    ```bash
    python start_conversion.py
    ```
*   The script will process all `.csv` files in `input/` and save the resulting `.xls` files in the `output/` directory.

**2. Single File Mode:**

*   Use the `-i` (or `--input`) flag for the input CSV path and the `-o` (or `--output`) flag for the desired output XLS path.
    ```bash
    python start_conversion.py -i path/to/your/data.csv -o path/to/your/output.xls
    ```
    *Example using default folders:*
    ```bash
    python start_conversion.py -i input/specific_file.csv -o output/converted_specific_file.xls
    ```

## Notes

*   The script assumes input CSV files are UTF-8 encoded.
*   The `output/` directory will be created automatically if it doesn't exist when running in folder mode.
*   Error messages and progress information will be printed to the console.
*   The output format is the older `.xls` (Excel 97-2003). This format has limitations, such as a maximum of 65,536 rows per sheet.
