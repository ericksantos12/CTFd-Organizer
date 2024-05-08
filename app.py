from os import mkdir, path

import pandas as pd
from xlsxwriter import Workbook
from rich.console import Console
from rich.prompt import Confirm
from rich.traceback import install
install()

console = Console()

package = {}

def start():
    try:
        input()
    except Exception as e:
        console.print('[red bold]Error:', e)
        exit()
        
    reader()
    merger()
    solve_sorter()
    excel_writer()
    
def input():
    """
    This function prompts the user for input and sets the paths for the CSV files.
    """

    console.print('[red bold]Make sure the files are in "./csv" folder')
    
    # Set the paths for the CSV files
    challenges_path = path.abspath('./csv/challenges.csv')
    users_path = path.abspath('./csv/users.csv')
    solves_path = path.abspath('./csv/solves.csv')

    # Confirm if the user wants to continue
    confirm = Confirm.ask('Do you want to continue? Output file will be overwritten', default=True)
    
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
    package.update({'challenges_path': challenges_path, 'users_path': users_path, 'solves_path': solves_path})

def reader():
    """
    Read CSV files and store the data in Dataframes.
    """

    # Reading CSV into Dataframe
    console.print('[yellow][Reader][/] Reading CSV files')
    df_solves = pd.read_csv(package.get('solves_path'), usecols=['challenge_id', 'user_id', 'type', 'date'])
    df_challenges = pd.read_csv(package.get('challenges_path'), usecols=['id', 'name'])
    df_users = pd.read_csv(package.get('users_path'), usecols=['id', 'name', 'affiliation'])
    
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
    
    # Close the writer object to save the Excel file
    writer.close()

if __name__ == '__main__':
    start()
    console.print('[green]Done[/]')