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
