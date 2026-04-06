import pandas as pd
import os
from utils import log

current_file_name = "file_handling.py"

def file_already_exists(file_name: str) -> bool:
    """Checks if the specified file exists

    Args:
        file_name (str): _description_

    Returns:
        bool: file existance
    """
    return os.path.exists(file_name)

def create_excel_file(file_name: str, columns: list[str]) -> bool:
    """Creates blank excel file

    Args:
        file_name (str): name of the file
        columns (list[str]): list of column names

    Returns:
        bool: acknowledgement
    """
    if file_already_exists(file_name):
        return True
    
    try:
        blank_data = {}
        for col_name in columns:
            blank_data[col_name] = []

        df = pd.DataFrame(blank_data, columns = columns)
        df = df[columns] # avoiding 'Unnamed: 0' column
        df.to_excel(os.path.join(os.path.dirname(__file__), file_name), index = False)
    except Exception as err:
        log(f"{current_file_name} {err}")
        return False
    return True