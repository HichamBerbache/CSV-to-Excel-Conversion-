import csv
import openpyxl
import os
import argparse

def convert_csv_to_excel(csv_name, sep, excel_name, sheet_name):
    # opening the files
    try:
        wb = openpyxl.load_workbook(excel_name)
        sheet = wb.get_sheet_by_name(sheet_name)
    except FileNotFoundError:
        raise Exception(f"The file {excel_name} or sheet {sheet_name} does not exist.")
    except Exception as e:
        raise Exception(f"An error occurred while opening the Excel file: {e}")

    try:
        file = open(csv_name, "r", encoding="utf-8")
    except FileNotFoundError:
        raise Exception(f"The file {csv_name} does not exist.")
    except Exception as e:
        raise Exception(f"An error occurred while opening the CSV file: {e}")

    # rows and columns
    row = 1
    column = 1

    # for each line in the file
    reader = csv.reader(file, delimiter=sep)
    for line in reader:
        # for each data in the line
        for data in line:
            # write the data to the cell
            sheet.cell(row, column).value = data
            # after each data column number increases by 1
            column += 1

        # to write the next line column number is set to 1 and row number is increased by 1
        column = 1
        row += 1

    # saving the excel file and closing the csv file
    wb.save(excel_name)
    file.close()

def main():
    # Set up command line arguments
    parser = argparse.ArgumentParser(description="Convert a CSV file to an Excel file.")
    parser.add_argument("input_file", help="Name of the CSV file for input (with the extension)")
    parser.add_argument("sep", help="Separator of the CSV file")
    parser.add_argument("output_file", help="Name of the excel file for output (with the extension)")
    parser.add_argument("sheet_name", help="Name of the excel sheet for output")
    args = parser.parse_args()

    # Check if input and output files are in the same directory as the script
    script_dir = os.path.dirname(os.path.realpath(__file__))
    input_path = os.path.join(script_dir, args.input_file)
    output_path = os.path.join(script_dir, args.output_file)
    if not os.path.exists(input_path):
        raise Exception(f"The input file {args.input_file} is not in the same directory as the script.")
    if not os.path.exists(output_path):
        raise Exception(f"The output file {args.output_file} is not in the same directory as the script.")

    convert_csv_to_excel(args.input_file, args.sep, args.output_file, args.sheet_name)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"An error occurred: {e}")

