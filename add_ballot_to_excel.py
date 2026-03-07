import pandas as pd
import argparse
import os
from openpyxl import load_workbook


def find_cell_by_value(file_name, sheet_name, target_value):
    ## Load specified workbook and get the specified sheet
    ##  - NOTE: Use read_only=True for performance if only reading
    wb = load_workbook(file_name, read_only=True)
    ws = wb[sheet_name]

    found_cells = []
    
    ## Iterate over all rows in current worksheet
    for row in ws.iter_rows():
        for cell in row:
            ## Check if the cell's value matches the target value
            if cell.value == target_value:
                found_cells.append(cell)

    if found_cells:
        print(f"Found '{target_value}' in the following cells:")
        for cell in found_cells:
            print(f"- Cell coordinate: {cell.coordinate}, Row: {cell.row}, Column: {cell.column}")
            # Example: print the value from the second column (index 1 in 0-based list) of the same row
            # Note: with read_only=True and iter_rows, accessing cells by index is easier
            # Example to get another value if needed can be handled in a separate iteration if using read_only
            # for full functionality or use the below approach if not using read_only
            # print(f"  Value in column 2 of this row: {ws.cell(row=cell.row, column=2).value}") # This requires not using read_only=True
            
    else:
        print(f"Value '{target_value}' not found in the sheet.")
        
    wb.close()

    return found_cells[0].row if found_cells else None


def create_sheet(file, new_sheet_name):
    ## 1. Load the ballot into a DataFrame
    df = pd.read_csv('indieheads_ult_26.txt', 
                     sep='\r', 
                     header=None
                )

    ## 2. Use Pandas ExcelWriter to append the DataFrame as a new sheet
    try:
        with pd.ExcelWriter(file,
                            mode='a',  ## Open in append mode
                            engine='openpyxl', 
                            if_sheet_exists='replace'
                        ) as writer:
                df.to_excel(writer,
                            sheet_name=new_sheet_name,
                            index=False, 
                            header=False
                        )
                
        print(f'DataFrame successfully written to sheet {new_sheet_name} in {file}')

    except FileNotFoundError:
        print(f"Error: The file '{file}' was not found. Please ensure the file exists.")
    except Exception as e:
        print(f"An error occurred: {e}")

    
    # ## Find where the ballot ENDs in order to start the statisics portion
    # find_cell_by_value(file, new_sheet_name, 'END')


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('file', help='Rates Excel file path')
    parser.add_argument('new_sheet_name', help='Name for new sheet for the ballot')
    args = parser.parse_args()
    
    create_sheet(args.file, args.new_sheet_name)
