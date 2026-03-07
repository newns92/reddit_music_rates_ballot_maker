import pandas as pd
import argparse
import os
from openpyxl import load_workbook


from openpyxl import load_workbook

def find_cell_by_value(filename, sheet_name, search_value_bonus, search_value_end):
    ## Load the workbook (use read_only=True for performance if only reading)
    wb = load_workbook(filename, read_only=True)
    ws = wb[sheet_name]

    found_end_cells = []
    
    ## Iterate over all rows in the worksheet
    for row in ws.iter_rows():
        for cell in row:
            ## Check if the cell's value is "END"
            if cell.value == search_value_end:
                found_end_cells.append(cell)
            ## If not, check for "BONUS TRACKS" and save that cell, if needed
            elif cell.value == search_value_bonus:
                bonus_cell = cell

    if found_end_cells:
        print(f"Found '{search_value_end}' in the following cells:")
        for cell in found_end_cells:
            print(f"- END Cell coordinate: {cell.coordinate}, Row: {cell.row}, Column: {cell.column}")
            # Example: print the value from the second column (index 1 in 0-based list) of the same row
            # Note: with read_only=True and iter_rows, accessing cells by index is easier
            # Example to get another value if needed can be handled in a separate iteration if using read_only
            # for full functionality or use the below approach if not using read_only
            # print(f"  Value in column 2 of this row: {ws.cell(row=cell.row, column=2).value}") # This requires not using read_only=True
    else:
        print(f"Value '{search_value_end}' not found in the sheet.")

    if bonus_cell:
        print(f"Found '{search_value_bonus}' in the following cells:")
        print(f"- BONUS TRACKS Cell coordinate: {bonus_cell.coordinate}, Row: {bonus_cell.row}, Column: {bonus_cell.column}")
        # Example: print the value from the second column (index 1 in 0-based list) of the same row
        # Note: with read_only=True and iter_rows, accessing cells by index is easier
        # Example to get another value if needed can be handled in a separate iteration if using read_only
        # for full functionality or use the below approach if not using read_only
        # print(f"  Value in column 2 of this row: {ws.cell(row=cell.row, column=2).value}") # This requires not using read_only=True
    else:
        print(f"Value '{search_value_bonus}' not found in the sheet.")
        
    wb.close()

    # print(found_cells)
    
    # ## Get the BONUS TRACKS cell row number, if found
    # final_cells = []
    # # final_cells.append(bonus_cell.row if bonus_cell else None)
    # final_cells.append(bonus_cell if bonus_cell else None)

    # ## Then, get the max row number from the found END cell(s) to get to bottom of ballot
    # ##  - i.e., After a bonus bonus rate, if exists
    # # return found_end_cells[-1].row if found_end_cells else None
    # # final_cells.append(found_end_cells[-1].row if found_end_cells else None)
    # final_cells.append(found_end_cells[-1] if found_end_cells else None)

    ## Get the BONUS TRACKS cell row number, if found
    final_cells = {}
    # final_cells.append(bonus_cell.row if bonus_cell else None)
    # final_cells.append(bonus_cell if bonus_cell else None)
    final_cells['bonus'] = bonus_cell if bonus_cell else None

    ## Then, get the max row number from the found END cell(s) to get to bottom of ballot
    ##  - i.e., After a bonus bonus rate, if exists
    # return found_end_cells[-1].row if found_end_cells else None
    # final_cells.append(found_end_cells[-1].row if found_end_cells else None)
    # final_cells.append(found_end_cells[-1] if found_end_cells else None)
    final_cells['end'] = found_end_cells[-1] if found_end_cells else None

    return final_cells


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

    ## 3. Find where alsbums end and ballot ENDs in order to start the statisics portion
    final_cells = find_cell_by_value(file, new_sheet_name, 'BONUS TRACKS', 'END')
    print(final_cells)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('filepath', help='Rates Excel file path')
    parser.add_argument('new_sheet_name', help='Name for new sheet for the ballot')
    args = parser.parse_args()
    
    create_sheet(args.filepath, args.new_sheet_name)
