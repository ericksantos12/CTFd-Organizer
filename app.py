from os import mkdir, path

import pandas as pd
from rich.console import Console
from rich.prompt import Confirm, Prompt

console = Console()

package = {}

def input():
    console.print('[yellow bold]Enter the corresponding path')
    
    challenges_path = path.abspath(Prompt.ask('[blue]Challenges[/] CSV file'))
    users_path = path.abspath(Prompt.ask('[blue]Users[/] CSV file'))
    solves_path = path.abspath(Prompt.ask('[blue]Solves[/] CSV file'))
    
    confirm = Confirm.ask('Do you want to continue? Output file will be overwritten', default=True)
    
    if confirm == False:
        exit()
    
    package.update({'challenges_path': challenges_path, 'users_path': users_path, 'solves_path': solves_path})

def reader():
    # Reading CSV into Dataframe
    console.print('[yellow][Reader][/] Reading CSV files')
    df_solves = pd.read_csv(package.get('solves_path'), usecols=['challenge_id', 'user_id', 'type'])
    df_challenges = pd.read_csv(package.get('challenges_path'), usecols=['id', 'name'])
    df_users = pd.read_csv(package.get('users_path'), usecols=['id', 'name'])
    
    # Renaming Columns
    df_challenges = df_challenges.rename(columns={'id': 'challenge_id', 'name': 'challenge_name'})
    df_users = df_users.rename(columns={'id': 'user_id', 'name': 'user_name'})
    
    
    # Packing dataframes
    package.update({'solves': df_solves, 'challenges': df_challenges, 'users': df_users})
