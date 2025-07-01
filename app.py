from os import mkdir, path

import pandas as pd
from xlsxwriter import Workbook
from rich.console import Console
from rich.prompt import Confirm, Prompt
from rich.traceback import install
from xlsxwriter.utility import xl_col_to_name
install()

console = Console()

package = {}

def start():
    try:
        get_user_input()
    except Exception as e:
        console.print('[red bold]Error:', e)
        exit()
        
    reader()
    merger()
    solve_sorter()
    excel_writer()
    
def get_user_input():
    """
    This function prompts the user for input and sets the paths for the CSV files.
    """

    console.print('[red bold]Make sure the files are in "./csv" folder')
    
    # Set the paths for the CSV files
    challenges_path = path.abspath('./csv/challenges.csv')
    users_path = path.abspath('./csv/users.csv')
    solves_path = path.abspath('./csv/solves.csv')

    challenge_filter = Prompt.ask('Get only visible (v), hidden (h), or all (a) challenges?', choices=["v", "h", "a"], default="a")

    # Confirm if the user wants to continue
    Confirm.ask('Do you want to continue? Output file will be overwritten', default=True)
    
    # Check if the challenges file exists
    if not path.exists(challenges_path):
        raise FileNotFoundError('Challenges not found')
    
    # Check if the users file exists
    if not path.exists(users_path):
        raise FileNotFoundError('Users not found')
    
    # Check if the solves file exists
    if not path.exists(solves_path):
        raise FileNotFoundError('Solves not found')
            
    console.print("[green]Starting...[/]")
    package.update({'challenges_path': challenges_path, 'users_path': users_path, 'solves_path': solves_path, 'challenge_filter': challenge_filter})

def reader():
    """
    Read CSV files and store the data in Dataframes.
    """

    # Reading CSV into Dataframe
    console.print('[yellow][Reader][/] Reading CSV files')
    df_solves = pd.read_csv(package.get('solves_path'), usecols=['challenge_id', 'user_id', 'type', 'date'])
    df_challenges = pd.read_csv(package.get('challenges_path'), usecols=['id', 'name', 'state'])
    df_users = pd.read_csv(package.get('users_path'), usecols=['id', 'name', 'affiliation'])
    
    # Filter challenges based on challenge_filter
    challenge_filter = package.get('challenge_filter')
    if challenge_filter == 'h':
        df_challenges = df_challenges[df_challenges['state'] == 'hidden']
        console.print('[yellow][Reader][/] Filtered to show only hidden challenges')
    elif challenge_filter == 'v':
        df_challenges = df_challenges[df_challenges['state'] == 'visible']
        console.print('[yellow][Reader][/] Filtered to show only visible challenges')
    
    # Renaming Columns
    df_challenges = df_challenges.rename(columns={'id': 'challenge_id', 'name': 'challenge_name'})
    df_users = df_users.rename(columns={'id': 'user_id', 'name': 'user_name'})
    
    # Packing dataframes
    package.update({'solves': df_solves, 'challenges': df_challenges, 'users': df_users})

def merger():
    """
    This function merges DataFrames, drops columns, replaces values,
    pivots DataFrames, reindexes columns, and updates the package.
    """

    # Merging Dataframes
    console.print('[yellow][Merger][/] Merging Dataframes')
    df_merged = pd.merge(package.get('solves'), package.get('challenges'), on='challenge_id')
    df_merged = pd.merge(df_merged, package.get('users'), how='right', on='user_id')

    # Dropping column ids
    df_merged.drop(['challenge_id', 'user_id'], axis=1, inplace=True)

    # Replacing values
    df_merged = df_merged.replace('correct', True)

    # Pivotting Dataframes
    console.print('[yellow][Merger][/] Pivotting Dataframes')
    try:
        with pd.option_context('future.no_silent_downcasting', True):
            df_pivot = df_merged.pivot(values='type', index=['user_name', 'affiliation'], columns='challenge_name').fillna(False)
            
    except ValueError as e:
        console.print('[red bold]Error:', "Duplicate values found in the csv files")
        exit()

    # Reindexing Columns
    change_columns = package.get('challenges')['challenge_name'].values.tolist()
    df_pivot = df_pivot.reindex(columns=change_columns)
    
    # Renaming index levels
    df_pivot = df_pivot.rename_axis(['RM', 'Nome'])

    # Packing in reverse order
    package.update({'pivot': df_pivot[::-1]})
    
def solve_sorter():
    """
    Sorts the 'solves' dataframe by date and updates the package with the sorted dataframe.
    """
    # Print a message indicating the sorting process has started
    console.print('[yellow][Sorter][/] Sorting Dataframes by date')
    
    # Merge the 'solves' dataframe with the 'challenges' dataframe based on the 'challenge_id' column
    df = pd.merge(package.get('solves'), package.get('challenges'), on='challenge_id')
    
    # Merge the resulting dataframe with the 'users' dataframe based on the 'user_id' column
    df = pd.merge(df, package.get('users'), on='user_id')
        
    # Convert the 'date' column to datetime format
    df["date"] = pd.to_datetime(df["date"], format='mixed')
    
    # Drop unnecessary columns from the dataframe
    df.drop(['challenge_id', 'user_id', 'type'], axis=1, inplace=True)
    
    # Reset the index of the dataframe
    df = df.reset_index(drop=True)
       
    # Select specific columns and sort the dataframe by the 'date' column
    columns = ['challenge_name', 'affiliation', 'user_name', 'date']
    df = df[columns].sort_values(by='date')
    
    # Update the package with the sorted dataframe
    package.update({'sorted_solves': df})
    
def excel_writer():
    """
    Writes data to an Excel file.
    """
    
    # Print a message indicating that we are writing to Excel
    console.print('[yellow][Writer][/] Writing to Excel')
    
    # Set the output path to the absolute path of the 'result' directory
    output_path = path.abspath('./result/')
    
    # If the 'result' directory does not exist, create it
    if not path.exists(output_path):
        mkdir(output_path)
        
    # Create an Excel writer object with the output file path and the 'xlsxwriter' engine
    writer = pd.ExcelWriter(path.join(output_path, 'output.xlsx'), engine='xlsxwriter')
    
    # Write the 'pivot' dataframe to the 'pivot' sheet of the Excel file
    package.get('pivot').to_excel(writer, sheet_name='pivot')
    
    # Write the 'sorted_solves' dataframe to the 'sorted_solves' sheet of the Excel file, without including the index
    package.get('sorted_solves').to_excel(writer, sheet_name='sorted_solves', index=False)
    
    # Print a message indicating that we are formatting the spreadsheet
    console.print('[yellow][Writer][/] Pretty formatting')
    
    workbook = writer.book
    worksheet_main = writer.sheets['pivot']
    
    # Get the dimensions of the dataframe
    (num_rows, num_cols) = package.get('pivot').shape
    
    # Get the headers from the dataframe and add 'Nota' and 'Peso'
    headers = [{'header': name} for name in package.get('pivot').reset_index().columns] + [{'header': 'Nota'}, {'header': 'Peso'}]
    
    # Create a table including 'Nota' and 'Peso' columns
    worksheet_main.add_table(0, 0, num_rows, num_cols + 3, {'columns': headers})
    
    # Autofit columns for 'pivot' sheet
    df_pivot_reset = package.get('pivot').reset_index()
    for i, col in enumerate(df_pivot_reset.columns):
        # Find the maximum length of the column header and the data
        column_len = max(df_pivot_reset[col].astype(str).map(len).max(), len(col))
        # Set the column width with a little padding
        worksheet_main.set_column(i, i, column_len + 10)

    # Create formats
    header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    white_format = workbook.add_format({'bg_color': 'white', 'align': 'center', 'valign': 'vcenter'})
    white_format_border = workbook.add_format({'bg_color': 'white', 'align': 'center', 'valign': 'vcenter', 'top': 1, 'bottom': 1, 'top_color': 'blue', 'bottom_color': 'blue'})

    # Center and bold the header row
    worksheet_main.set_row(0, None, header_format)

    # Apply white_format only to the data rows inside the table (skip header)
    worksheet_main.conditional_format(
        1, 0,
        num_rows, 1,
        {
            'type':      'no_errors',
            'format':    white_format_border
        }
    )
    
    # Define custom formats matching Excel's "Good" and "Bad" styles
    good_fmt = workbook.add_format({
        'bg_color': '#C6EFCE',
        'font_color': '#006100',
        'align': 'center',
        'valign': 'vcenter',
        'top': 1, 
        'bottom': 1, 
        'top_color': 'blue', 
        'bottom_color': 'blue'
    })
    bad_fmt = workbook.add_format({
        'bg_color': '#FFC7CE',
        'font_color': '#9C0006',
        'align': 'center',
        'valign': 'vcenter',
        'top': 1, 
        'bottom': 1, 
        'top_color': 'blue', 
        'bottom_color': 'blue'
    })

    # Apply conditional formatting from the 3rd column onward
    start_row, start_col = 1, 2
    end_row, end_col = num_rows, num_cols + 1

    # VERDADEIRO → Good style
    worksheet_main.conditional_format(
        start_row, start_col, end_row, end_col,
        {
            'type': 'text',
            'criteria': 'containing',
            'value': 'VERDADEIRO',
            'format': good_fmt
        }
    )

    # FALSO → Bad style
    worksheet_main.conditional_format(
        start_row, start_col, end_row, end_col,
        {
            'type': 'text',
            'criteria': 'containing',
            'value': 'FALSO',
            'format': bad_fmt
        }
    )
    
    # Determine the columns where we'll write Nota and Peso
    nota_col = num_cols + 2     # zero‐based index of “Nota”
    peso_col = num_cols + 3     # zero‐based index of “Peso”

    # Write the headers for the two new columns (already defined in the table)
    worksheet_main.write(0, nota_col, 'Nota', header_format)
    worksheet_main.write(0, peso_col, 'Peso', header_format)
    
    worksheet_main.conditional_format(
        1, nota_col,
        num_rows, peso_col,
        {
            'type':      'no_errors',
            'format':    white_format
        }
    )

    # Put the Peso value in the first data row (only one row)
    peso_val = 20
    worksheet_main.write(1, peso_col, peso_val)
    
    # Build a COUNTIF formula for each row to compute Nota = (count of VERDADEIRO)*Peso/20
    first_data_col = 2                # first challenge column (zero‐based)
    last_data_col = num_cols + 1      # last challenge column (zero‐based)
    for row in range(1, num_rows + 1):
        # convert to A1‐style
        start_cell = xl_col_to_name(first_data_col) + str(row + 1)
        end_cell   = xl_col_to_name(last_data_col)   + str(row + 1)
        peso_ref   = xl_col_to_name(peso_col)        + '$2'
        formula = (
            f"=COUNTIF({start_cell}:{end_cell},\"VERDADEIRO\")"
            f"*{peso_ref}/20"
        )
        worksheet_main.write_formula(row, nota_col, formula)

    # Apply a data‐bar conditional format to the “Nota” column
    worksheet_main.conditional_format(
        1, nota_col, num_rows, nota_col,
        {
            'type': 'data_bar',
            'bar_color': '#008AEF'
        }
    )
    
    writer.close()


if __name__ == '__main__':
    start()
    console.print('[green]Done[/]')