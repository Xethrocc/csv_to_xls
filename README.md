# CSV to XLSX Converter

This script converts CSV (Comma Separated Values) files into XLSX (Excel) format.

It can operate in two modes:

1.  **Folder Mode:** Converts all `.csv` files found in a specified input directory (default: `input/`) and saves the corresponding `.xlsx` files to an output directory (default: `output/`).
2.  **Single File Mode:** Converts a specific input CSV file to a specified output XLSX file.

The script automatically tries to detect whether the CSV delimiter is a comma (`,`) or a semicolon (`;`).

## Project Structure

```
.
├── csv_to_xlsx_converter/ # Contains the main conversion logic
│   └── converter.py
├── input/                 # Default directory for input CSV files
├── output/                # Default directory for output XLSX files (created if needed)
├── venv/                  # Example Python virtual environment (ignored by Git)
├── .gitignore             # Specifies files/folders for Git to ignore
├── README.md              # This file
├── requirements.txt       # (Optional but recommended) Lists dependencies
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
    The script requires the `openpyxl` library. You can install it directly or (preferably) from a `requirements.txt` file if one is provided.
    ```bash
    pip install openpyxl
    ```
    *If a `requirements.txt` exists:* 
    ```bash
    pip install -r requirements.txt
    ```

4.  **Prepare Input Files:**
    *   Place your `.csv` files into the `input/` directory (or create it if it doesn't exist).

## Usage

You can run the conversion using the `start_conversion.py` script from the project's root directory.

**1. Folder Mode (Default):**

*   Ensure your CSV files are in the `input/` directory.
*   Run the start script:
    ```bash
    python start_conversion.py
    ```
*   The script will process all `.csv` files in `input/` and save the resulting `.xlsx` files in the `output/` directory.

**2. Single File Mode:**

*   Use the `-i` (or `--input`) flag for the input CSV path and the `-o` (or `--output`) flag for the desired output XLSX path.
    ```bash
    python start_conversion.py -i path/to/your/data.csv -o path/to/your/output.xlsx
    ```
    *Example using default folders:*
    ```bash
    python start_conversion.py -i input/specific_file.csv -o output/converted_specific_file.xlsx
    ```

## Notes

*   The script assumes input CSV files are UTF-8 encoded.
*   The `output/` directory will be created automatically if it doesn't exist when running in folder mode.
*   Error messages and progress information will be printed to the console.
