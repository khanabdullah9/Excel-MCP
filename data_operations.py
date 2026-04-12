import pandas as pd
from file_handling import *
import xlwings as xw

def get_data(file_name: str, sheet_name: str) -> pd.DataFrame:
    """Retrieves data from the excel file

    Args:
        file_name (str): _description_

    Returns:
        pd.DataFrame: dataset
    """
    if not file_already_exists(file_name):
        return pd.DataFrame({})
    
    return pd.read_excel(file_name, sheet_name = sheet_name)

def data_preprocess(file_name: str, sheet_name: str, data: dict) -> tuple[pd.DataFrame, str]:
    """Removes unncessary data

    Args:
        file_name (str): _description_
        sheet_name (str): name fo the sheet in the file
        data (dict): raw user data

    Returns:
        tuple[pd.DataFrame, str]: (updated df, write mode -> append/write)
    """
    if not file_already_exists(file_name):
        create_excel_file(file_name=file_name, columns=list(data.keys()))

    df_new = pd.DataFrame(data)
    df_new = df_new[list(data.keys())] # avoiding unncessary column inclusion
    # df_new = df_new.drop_duplicates()

    existing_data = get_data(file_name=file_name, sheet_name = sheet_name)
    if existing_data.empty:
        return df_new

    merged = df_new.merge(
        existing_data, 
        on=list(data.keys()), 
        how='left', 
        indicator=True
    )

    df_final = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge'])

    return df_final

def update_row(row: pd.Series, data: dict) -> pd.Series:
    new_row = row.copy()

    for col in row.index:
        if col in data:
            new_row[col] = data[col]

    return new_row


    