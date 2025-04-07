import csv
import sys
import xlwt # Changed from openpyxl
import os
import argparse

def read_csv_data(file_path):
    """Reads data from a CSV file, attempting to detect comma or semicolon delimiters.

    Args:
        file_path (str): The path to the CSV file.

    Returns:
        list: A list of lists representing the rows and columns of the CSV data.
              Returns an empty list if the file cannot be read or delimiter is unclear.
    """
    data = []
    try:
        with open(file_path, mode='r', newline='', encoding='utf-8') as csvfile:
            # --- Delimiter Sniffing ---
            try:
                # Read a sample for sniffing, then reset pointer
                sample = csvfile.read(2048) # Read up to 2KB
                dialect = csv.Sniffer().sniff(sample, delimiters=',;')
                csvfile.seek(0) # Go back to the start of the file
                delimiter = dialect.delimiter
                print(f"Detected delimiter: '{delimiter}'")
            except csv.Error:
                # If sniffing fails (e.g., empty file, ambiguous), default to comma
                csvfile.seek(0) # Go back to the start of the file
                delimiter = ','
                print("Warning: Could not automatically detect delimiter. Assuming ','.", file=sys.stderr)
            # --------------------------

            reader = csv.reader(csvfile, delimiter=delimiter)
            for row in reader:
                # Handle potential empty rows often found in CSVs
                if row: # Only append if row is not empty
                    data.append(row)
        if not data:
             print(f"Warning: No data read from {file_path}. File might be empty or incorrectly formatted.", file=sys.stderr)
        else:
            print(f"Successfully read data from {file_path}")
        return data
    except FileNotFoundError:
        print(f"Error: File not found at {file_path}", file=sys.stderr)
        return [] # Return empty list on error
    except Exception as e:
        print(f"Error reading CSV file {file_path}: {e}", file=sys.stderr)
        return [] # Return empty list on other errors

def write_xls_data(data, output_file_path): # Renamed function
    """Writes data to an XLS file using xlwt.

    Args:
        data (list): A list of lists representing the data to write.
        output_file_path (str): The path to the output XLS file.

    Returns:
        bool: True if writing was successful, False otherwise.
    """
    try:
        workbook = xlwt.Workbook() # Use xlwt Workbook
        sheet = workbook.add_sheet('Sheet1') # Add a sheet

        for r_idx, row_data in enumerate(data):
            for c_idx, item in enumerate(row_data):
                # Attempt to convert numeric strings to numbers
                try:
                    # Try converting to float first (handles ints and floats)
                    numeric_value = float(item)
                    sheet.write(r_idx, c_idx, numeric_value) # Write numeric value
                except (ValueError, TypeError):
                    # If conversion fails, write it as a string
                    sheet.write(r_idx, c_idx, item) # Write string value

        workbook.save(output_file_path) # Save the workbook
        print(f"Successfully wrote data to {output_file_path}")
        return True
    except PermissionError:
        print(f"Error: Permission denied when trying to write to {output_file_path}", file=sys.stderr)
        return False
    except Exception as e:
        print(f"Error writing XLS file {output_file_path}: {e}", file=sys.stderr)
        return False


# --- Function for single file conversion ---
def process_single_file(input_path, output_path):
    """Handles the conversion of a single CSV file to XLS.""" # Updated docstring
    print(f"--- Processing single file ---")
    print(f"Input: {input_path}")
    print(f"Output: {output_path}")

    # --- Input File Validation ---
    if not os.path.isfile(input_path):
        print(f"Error: Input file not found at {input_path}", file=sys.stderr)
        sys.exit(1)
    # ---------------------------

    csv_data = read_csv_data(input_path)

    if csv_data:
        print("CSV data read successfully.")
        if write_xls_data(csv_data, output_path): # Call renamed function
            print("Conversion completed successfully.")
        else:
            print("Conversion failed during XLS writing.") # Updated message
            sys.exit(1) # Exit with an error code
    else:
        print(f"Could not read CSV data from {input_path}. Exiting.")
        sys.exit(1) # Exit with an error code

# --- Function for folder processing ---
def process_folder(input_dir='input', output_dir='output'):
    """Handles converting all CSV files in an input folder to an output folder."""
    print(f"--- Processing folder ---")
    print(f"Input directory: {input_dir}")
    print(f"Output directory: {output_dir}")

    # --- Directory Validation and Creation ---
    if not os.path.isdir(input_dir):
        print(f"Error: Input directory '{input_dir}' not found.", file=sys.stderr)
        sys.exit(1)

    try:
        os.makedirs(output_dir, exist_ok=True)
        print(f"Ensured output directory '{output_dir}' exists.")
    except OSError as e:
        print(f"Error: Could not create output directory '{output_dir}': {e}", file=sys.stderr)
        sys.exit(1)
    # ---------------------------------------

    converted_count = 0
    skipped_count = 0
    # --- File Iteration ---
    print("Scanning for CSV files...")
    for filename in os.listdir(input_dir):
        if filename.lower().endswith('.csv'):
            input_path = os.path.join(input_dir, filename)
            # Construct output path, replacing .csv with .xls (case-insensitive)
            base, _ = os.path.splitext(filename)
            output_filename = base + '.xls' # Changed extension to .xls
            output_path = os.path.join(output_dir, output_filename)

            print(f"\nProcessing '{filename}' -> '{output_filename}'")

            csv_data = read_csv_data(input_path)
            if csv_data:
                if write_xls_data(csv_data, output_path): # Call renamed function
                    print(f"Successfully converted '{filename}'.")
                    converted_count += 1
                else:
                    print(f"Failed to write XLS for '{filename}'. Skipping.") # Updated message
                    skipped_count += 1
            else:
                print(f"Could not read data from '{filename}'. Skipping.")
                skipped_count += 1
        else:
            # Optional: print message for non-csv files
            # print(f"Skipping non-CSV file: {filename}")
            pass
    # ---------------------

    print(f"\n--- Folder processing summary ---")
    print(f"Successfully converted: {converted_count} file(s)")
    print(f"Skipped/Failed:       {skipped_count} file(s)")
    print(f"---------------------------------")

# --- Main execution logic ---
def main():
    parser = argparse.ArgumentParser(
        description="Convert CSV file(s) to XLS format. Detects comma or semicolon delimiters.", # Updated description
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  To convert a single file:
    python converter.py -i data.csv -o converted_data.xls # Updated example extension

  To automatically convert all *.csv files in the 'input' folder to the 'output' folder (as *.xls): # Updated example description
    python converter.py
"""
    )
    parser.add_argument("-i", "--input", help="Path to the input CSV file (for single file conversion).", metavar='FILE')
    parser.add_argument("-o", "--output", help="Path for the output XLS file (for single file conversion).", metavar='FILE') # Updated help text

    args = parser.parse_args()

    # Decide mode based on arguments
    if args.input and args.output:
        # Single file mode
        # Ensure output file has .xls extension if user didn't specify
        if not args.output.lower().endswith('.xls'):
             base, _ = os.path.splitext(args.output)
             args.output = base + '.xls'
             print(f"Warning: Output filename did not end with .xls. Changed to: {args.output}", file=sys.stderr)
        process_single_file(args.input, args.output)
    elif not args.input and not args.output:
        # Folder processing mode (default folders: 'input' and 'output')
        process_folder()
    else:
        # Error: User provided only one of -i or -o OR invalid combination
        print("Error: For single file conversion, both -i/--input and -o/--output arguments are required.", file=sys.stderr)
        print("To use automatic folder processing, run the script without any arguments.", file=sys.stderr)
        parser.print_help(sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
