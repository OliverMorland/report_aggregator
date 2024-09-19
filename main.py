import os
from excel_utils import consolidate_excel_files
import sys


def get_arguments():
    # Check if at least two arguments are provided
    if len(sys.argv) < 3:
        print("Please provide at least two arguments.")
        sys.exit(1)

    # Get the first and second arguments
    first_arg = sys.argv[1]
    second_arg = sys.argv[2]

    return first_arg, second_arg


def main():

    # Get paths
    directory, output_file = get_arguments()

    # Directory where the Excel files are located
    # directory = os.path.join(os.path.dirname(__file__), 'data')

    # Output file for consolidated data
    # output_file = os.path.join(os.path.dirname(__file__), 'aggregated_data.xlsx')

    # Call the function to consolidate Excel files
    consolidate_excel_files(directory, output_file)

    print(f"Consolidation complete. File saved as: {output_file}")


if __name__ == "__main__":
    main()
