from os import mkdir, path

import pandas as pd
from xlsxwriter import Workbook
from rich.console import Console
from rich.prompt import Confirm, Prompt
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

    console.print('[red bold]Make sure the files are in "./csv" folder')
    
    challenges_path = path.abspath('./csv/challenges.csv')
    users_path = path.abspath('./csv/users.csv')
    solves_path = path.abspath('./csv/solves.csv')

    confirm = Confirm.ask('Do you want to continue? Output file will be overwritten', default=True)
    
    if not path.exists(challenges_path):
        raise FileNotFoundError('Challenges not found')
    
    if not path.exists(users_path):
        raise FileNotFoundError('Users not found')
    
    if not path.exists(solves_path):
        raise FileNotFoundError('Solves not found')
    console.print("[green]Starting...[/]")
    package.update({'challenges_path': challenges_path, 'users_path': users_path, 'solves_path': solves_path})

def reader():
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
    df_pivot = df_merged.pivot(values='type', index=['user_name', 'affiliation'], columns='challenge_name').fillna(False)
    
    
    # Reindexing Columns
    change_columns = package.get('challenges')['challenge_name'].values.tolist()
    df_pivot = df_pivot.reindex(columns=change_columns)
    
    # Packing in reverse order
    package.update({'pivot': df_pivot[::-1]})
    
def solve_sorter():
    console.print('[yellow][Sorter][/] Sorting Dataframes by date')
    
    df = pd.merge(package.get('solves'), package.get('challenges'), on='challenge_id')
    df = pd.merge(df, package.get('users'), on='user_id')
        
    df["date"] = pd.to_datetime(df["date"])
    
    df.drop(['challenge_id', 'user_id', 'type'], axis=1, inplace=True)
    df = df.reset_index(drop=True)
       
    columns = ['challenge_name', 'affiliation', 'user_name', 'date']
    
    df = df[columns].sort_values(by='date')
    
    package.update({'sorted_solves': df})
    
    
def excel_writer():
    console.print('[yellow][Writer][/] Writing to Excel')
    
    output_path = path.abspath('./result/')
    
    if not path.exists(output_path):
        mkdir(output_path)
        
    writer = pd.ExcelWriter(path.join(output_path, 'output.xlsx'), engine='xlsxwriter')
    
    package.get('pivot').to_excel(writer, sheet_name='pivot')
    package.get('sorted_solves').to_excel(writer, sheet_name='sorted_solves', index=False)
    writer.close()

if __name__ == '__main__':
    start()
    console.print('[green]Done[/]')