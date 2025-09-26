# CSV Comparison Utility

This utility compares two large CSV files by a specified key column and outputs the results to an Excel (.xlsx) file, automatically splitting data into multiple sheets if the row count exceeds Excel's per-sheet limit.

## File Structure
- `compare_csv_files.py`: Main entry point. Handles logging and calls the comparison logic.
- `comparator.py`: Contains the main comparison logic.
- `csv_indexer.py`: Functions for building CSV indexes and reading rows by offset.
- `excel_writer.py`: Functions for writing to Excel and handling sheet splitting.
- `file1.csv`, `file2.csv`: Example input CSV files.
- `comparison_output.xlsx`: Output Excel file (created after running the program).

## Input Parameters
- **file1**: Path to the first CSV file (default: `compare_utility/file1.csv`)
- **file2**: Path to the second CSV file (default: `compare_utility/file2.csv`)
- **output**: Path to the output Excel file (default: `compare_utility/comparison_output.xlsx`)
- **keycol**: Name of the key column to join on (default: `keycol`)

You can change these parameters in `compare_csv_files.py` or modify the script to accept command-line arguments.

## How to Execute
1. Open a terminal and navigate to the parent directory of `compare_utility` (e.g., `C:\Users\edura\PycharmProjects\python`).
2. Run the following command:

    ```sh
    python -m compare_utility.compare_csv_files
    ```

   This ensures Python treats `compare_utility` as a package and resolves all imports correctly.

## Expected Output
- An Excel file (`comparison_output.xlsx`) with the following sheets:
  - **Matching Rows**: Rows with matching keys in both files, with differences highlighted.
  - **Only in File1**: Rows unique to the first file.
  - **Only in File2**: Rows unique to the second file.
- If any sheet exceeds 1,048,576 rows, the output is split into multiple sheets (e.g., `Matching Rows_2`).
- The script logs progress and summary information to the console.

## Notes
- The script is optimized for large files and uses streaming Excel writing.
- Input CSVs must have a header row and the specified key column.
- The output Excel file is in `.xlsx` format, which is already compressed.

## Troubleshooting
- If you see import errors, ensure you are running the script from the parent directory and not from inside `compare_utility`.
- If you need to process different files or key columns, edit the variables in `compare_csv_files.py`.

