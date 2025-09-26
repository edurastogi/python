# CSV Comparison Utility

This utility compares two large CSV files by a specified key column and outputs the results to an Excel (.xlsx) file, automatically splitting data into multiple sheets if the row count exceeds Excel's per-sheet limit.

## Directory Structure
- `compare_csv_files.py`: Main entry point. Handles logging, input/output setup, and calls the comparison logic.
- `comparator.py`: Contains the main comparison logic and orchestrates the comparison process.
- `csv_indexer.py`: Functions for building CSV indexes and reading rows by offset.
- `excel_writer.py`: Functions for writing to Excel and handling sheet splitting.
- `input/`: Directory containing input CSV files (`file1.csv`, `file2.csv`).
- `output/`: Directory where output Excel files are saved (e.g., `comparison_output.xlsx`).
- `README.md`: This documentation file.

## Program Flow
1. **Start in `compare_csv_files.py`:**
   - Initializes logging.
   - Ensures input and output directories exist.
   - Sets file paths for input and output.
   - If the default output file exists, generates a new output filename with a timestamp.
   - Calls `compare_large_csv` from `comparator.py`.

2. **`compare_large_csv` in `comparator.py`:**
   - Builds key indexes and headers for both input CSVs using `build_key_index_and_header` from `csv_indexer.py`.
   - Checks that the key column exists in both files.
   - Determines all columns to be included in the output.
   - Creates an Excel workbook in write-only mode.
   - For each key:
     - If present in both files, compares and writes to the "Matching Rows" sheet (highlighting differences).
     - If only in file1, writes to the "Only in File1" sheet.
     - If only in file2, writes to the "Only in File2" sheet.
   - Uses `append_row_with_split` from `excel_writer.py` to handle sheet splitting if row limits are exceeded.
   - Saves the Excel file.

3. **`build_key_index_and_header` in `csv_indexer.py`:**
   - Reads the header from a CSV file.
   - Builds a dictionary mapping key column values to file offsets for efficient row access.

4. **`get_row_dict_by_offset` in `csv_indexer.py`:**
   - Given a file offset and header, retrieves a row as a dictionary from the CSV file.

5. **`append_row_with_split` in `excel_writer.py`:**
   - Appends a row to the current worksheet.
   - If the sheet reaches the Excel row limit, creates a new sheet with an incremented name and continues writing.

6. **`get_next_sheet_name` in `excel_writer.py`:**
   - Helper function to generate new sheet names with suffixes when splitting.

## Function Descriptions
- **main (compare_csv_files.py):** Sets up logging, directories, file paths, and calls the main comparison function.
- **compare_large_csv (comparator.py):** Orchestrates the entire comparison process and Excel output.
- **build_key_index_and_header (csv_indexer.py):** Builds a key-to-offset index and returns the header for a CSV file.
- **get_row_dict_by_offset (csv_indexer.py):** Retrieves a row as a dictionary from a CSV file given an offset and header.
- **append_row_with_split (excel_writer.py):** Appends a row to a worksheet, splitting to a new sheet if the row limit is reached.
- **get_next_sheet_name (excel_writer.py):** Generates a new sheet name with a suffix for split sheets.

## How to Execute
1. Place your input CSV files in the `compare_utility/input` directory. Ensure both files have a header row and the specified key column.
2. Open a terminal and navigate to the parent directory of `compare_utility` (e.g., `C:\Users\edura\PycharmProjects\python`).
3. Run the following command:

    ```sh
    python -m compare_utility.compare_csv_files
    ```

   This ensures Python treats `compare_utility` as a package and resolves all imports correctly.

## Expected Output
- An Excel file (e.g., `comparison_output.xlsx`) will be created in the `compare_utility/output` directory with the following sheets:
  - **Matching Rows**: Rows with matching keys in both files, with differences highlighted.
  - **Only in File1**: Rows unique to the first file.
  - **Only in File2**: Rows unique to the second file.
- If any sheet exceeds 1,048,576 rows, the output is split into multiple sheets (e.g., `Matching Rows_2`).
- If an output file with the default name already exists, a new file with a timestamp will be created automatically.
- The script logs progress and summary information to the console.

## Notes
- The script is optimized for large files and uses streaming Excel writing.
- Input CSVs must have a header row and the specified key column.
- The output Excel file is in `.xlsx` format, which is already compressed.

## Troubleshooting
- If you see import errors, ensure you are running the script from the parent directory and not from inside `compare_utility`.
- If you need to process different files or key columns, edit the variables in `compare_csv_files.py`.
