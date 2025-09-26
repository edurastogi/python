import logging
import os
import datetime
from compare_utility.comparator import compare_large_csv

def main():
    # Initialize logging before any log messages
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

    # Set up input and output directories
    input_dir = 'compare_utility/input'
    output_dir = 'compare_utility/output'
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    output_base = 'comparison_output.xlsx'
    output = os.path.join(output_dir, output_base)
    file1 = os.path.join(input_dir, 'file1.csv')
    file2 = os.path.join(input_dir, 'file2.csv')
    keycol = 'keycol'

    # Always use a new output file if the default exists
    if os.path.exists(output):
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        output = os.path.join(output_dir, f'comparison_output_{timestamp}.xlsx')
        logging.warning(f"Output file already exists. Using new output file: {output}")

    logging.info("Starting CSV comparison utility...")
    try:
        compare_large_csv(file1, file2, output, keycol)
        logging.info(f'Comparison complete. Output saved to {output}')
    except FileNotFoundError as e:
        logging.error(f"File not found: {e.filename}")
        print(f"Error: File not found: {e.filename}")
    except ValueError as e:
        logging.error(f"Value error: {e}")
        print(f"Error: {e}")
    except PermissionError as e:
        logging.error(f"Permission error: {e}")
        print(f"Error: Permission denied. {e}")
    except Exception as e:
        logging.critical(f"Unexpected error: {e}", exc_info=True)
        print(f"An unexpected error occurred: {e}")

if __name__ == '__main__':
    main()
