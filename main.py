import os
from excel_utils import consolidate_excel_files


def main():
    # Directory where the Excel files are located
    directory = os.path.join(os.path.dirname(__file__), 'data')

    # Output file for consolidated data
    output_file = os.path.join(os.path.dirname(__file__), 'aggregated_data.xlsx')

    # Call the function to consolidate Excel files
    consolidate_excel_files(directory, output_file)

    print(f"Consolidation complete. File saved as: {output_file}")


if __name__ == "__main__":
    main()
