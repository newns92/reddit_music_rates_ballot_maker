import pandas as pd
import os

def final_ballot(sheet_name, username):
    ## Make a directory to hold final ballot text files if it doesn't already exist
    if not os.path.exists('./ballots'):
        os.makedirs('./ballots')

    ## Read in data from provided sheet name arg, print out sheet names if not available
    try:
        df = pd.read_excel(f'./rates.xlsx', sheet_name=sheet_name, header=None)
        # print(df.head())

    except ValueError as e:
        print(f"Sheet '{sheet_name}' not found. Available sheets:\n")
        excel_file = pd.ExcelFile(f'./rates.xlsx')
        # print(excel_file.sheet_names)
        for sheet in excel_file.sheet_names:
            print(sheet, '\n')
        quit()

    ## Initiate list to hold results
    result_list = []

    ## For each row in dataframe...
    for idx, row in df.iterrows():
        ## Grab the value from the first column, stripping trailing whitespace
        first_col = str(row.iloc[0]).strip()
        
        ## If first row of sheet, just add "Username: " plus given username arg as a line
        if first_col.startswith("Username"):
            result_list.append(f'Username: {username}')
        ## If last row, just add "END" to not break rate machine
        elif first_col == "END":
            result_list.append(first_col)
            break
        ## Otherwise, if 'normal' row (song: | score | comment)...
        else:
            ## Concatenate first 3 columns (song, score, comment) with spaces,
            ##      whilst removing trailing space from first col
            cols = [str(row.iloc[i]) if pd.notna(row.iloc[i]) else '' for i in range(3)]
            result_list.append(cols[0].rstrip() + ' ' + ' '.join(cols[1:]))

    # print(result_list)

    ## TODO: Do NOT add in average from bonus tracks to the final ballot output

    ## Write results to corresponding text file
    with open(f'./ballots/{sheet_name}_rate_output.txt', 'w') as f:
        for line in result_list:
            if line == "END":
                f.write(line)
            ## Double-space for rate machine format
            else:
                f.write(line + '\n\n')

    print(f"\nResults written to {sheet_name}_rate_output.txt in {len(result_list)} lines")

if __name__ == "__main__":
    final_ballot('ServerRemix', 'nimz (Non-Vocal Mix)')
