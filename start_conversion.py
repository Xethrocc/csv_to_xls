import subprocess
import sys
import os

# Get the directory where this script is located (the project root)
project_root = os.path.dirname(os.path.abspath(__file__))
# Construct the path to the converter script
converter_script = os.path.join(project_root, 'src', 'converter.py')

# Check if the converter script exists
if not os.path.isfile(converter_script):
    print(f"Error: Converter script not found at {converter_script}", file=sys.stderr)
    sys.exit(1)

# --- Run the converter script ---
# We use sys.executable to ensure we're using the same Python interpreter
# that's running this start script (important if using a virtual environment).
# We pass any command-line arguments received by start_conversion.py directly
# to converter.py. This allows using -i and -o with the start script too.
command = [sys.executable, converter_script] + sys.argv[1:]

print(f"Running command: {' '.join(command)}")
try:
    # Run the script and wait for it to complete
    process = subprocess.run(command, check=True, text=True, capture_output=True)
    # Print the output (stdout and stderr) from the converter script
    if process.stdout:
        print("\n--- Converter Output ---")
        print(process.stdout)
    if process.stderr:
        print("\n--- Converter Errors ---", file=sys.stderr)
        print(process.stderr, file=sys.stderr)
    print("\n------------------------")
    print("Conversion process finished.")

except subprocess.CalledProcessError as e:
    # This catches errors if the converter script exits with a non-zero status
    print(f"Error during conversion process (Exit Code: {e.returncode}).", file=sys.stderr)
    if e.stdout:
        print("\n--- Converter Output ---")
        print(e.stdout)
    if e.stderr:
        print("\n--- Converter Errors ---", file=sys.stderr)
        print(e.stderr, file=sys.stderr)
    sys.exit(e.returncode) # Exit with the same error code as the converter
except Exception as e:
    print(f"An unexpected error occurred while trying to run the converter: {e}", file=sys.stderr)
    sys.exit(1)

