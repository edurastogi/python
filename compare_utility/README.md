# CSV Comparison Utility

This utility compares two large CSV files by a specified key column and outputs the results to an Excel (.xlsx) file, automatically splitting data into multiple sheets if the row count exceeds Excel's per-sheet limit.

## File Structure
- `compare_csv_files.py`: Main entry point. Handles logging and calls the comparison logic.
- `comparator.py`: Contains the main comparison logic.
- `csv_indexer.py`: Functions for building CSV indexes and reading rows by offset.
- `excel_writer.py`: Functions for writing to Excel and handling sheet splitting.
- `input/`: Directory containing input CSV files (`file1.csv`, `file2.csv`).
- `output/`: Directory where output Excel files are saved (e.g., `comparison_output.xlsx`).
- `README.md`: This documentation file.

## Input Parameters
- **file1**: Path to the first CSV file (default: `compare_utility/input/file1.csv`)
- **file2**: Path to the second CSV file (default: `compare_utility/input/file2.csv`)
- **output**: Path to the output Excel file (default: `compare_utility/output/comparison_output.xlsx`)
- **keycol**: Name of the key column to join on (default: `keycol`)

You can change these parameters in `compare_csv_files.py` or modify the script to accept command-line arguments.

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

## Refactoring & Improvements
- **Input and Output Folders:** Input files are now read from the `input` directory, and all Excel outputs are written to the `output` directory for better organization.
- **Automatic Output Naming:** If the output Excel file already exists (e.g., is open or locked), a new file with a timestamp is created to avoid errors.
- **Modular Codebase:** The code is split into logical modules for indexing, comparison, and Excel writing, making it easier to maintain and extend.
- **Robust Directory Handling:** The script ensures both input and output directories exist before running.

## Notes
- The script is optimized for large files and uses streaming Excel writing.
- Input CSVs must have a header row and the specified key column.
- The output Excel file is in `.xlsx` format, which is already compressed.

## Troubleshooting
- If you see import errors, ensure you are running the script from the parent directory and not from inside `compare_utility`.
- If you need to process different files or key columns, edit the variables in `compare_csv_files.py`.
