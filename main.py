from mcp.server.fastmcp import FastMCP
from data_operations import data_preprocess, update_row
import xlwings as xw
from utils import log
import pandas as pd
import os
from plot_charts import *
from file_handling import file_already_exists

mcp = FastMCP("Excel-Automation")
current_file_name = "main.py"

@mcp.tool()
@mcp.resource("file://{file_name}/{sheet_name}")
def get_table_schema(file_name: str, sheet_name: str = "Sheet1") -> list[str]:
    """Retrieves the schema of the table.
    Args:
        file_name (str): excel file name
        sheet_name (str, optional): excel sheet name. Defaults to "Sheet1".

    Returns:
        list[str]: list of column names
    """
    if not file_already_exists(file_name=file_name):
        return []

    try:
        wb = xw.Book(file_name)
        sheet = wb.sheets[sheet_name]

        # Instead of using .options(pd.DataFrame...)
        data = sheet.range("A1").expand("table").value

        if data:
            # Use the first row as columns, the rest as data
            df = pd.DataFrame(data[1:], columns=data[0])
            return df.columns.tolist()
        
    except Exception as err:
        log(f"{current_file_name} {err}")
    return []

@mcp.tool()
def write_data_live(file_name: str, data: dict, sheet_name: str = "Sheet1") -> bool:
    """Writes data to an Excel file even if it is currently open in Excel.
        The 'data' argument will be converted to a pandas dataframe.
            Required format: {
                        column_name: list[object]
                        }

    Args:
        file_name (str): Name of the output file
        data (dict): User's raw input data
        sheet_name (str, optional): Spreadsheet name. Defaults to "Sheet1"
    Returns:
        bool: Acknowledgement whether the process was successful
    """
    df_new = data_preprocess(file_name, sheet_name, data)

    try:
        wb = xw.Book(file_name)
        sheet = wb.sheets[sheet_name]

        last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end('up').row
        
        # if mode == "write":
        #     sheet.range("A1").options(index = True).value = df_new
        # else:
        #     sheet.range(f"A{last_row + 1}").options(index=True, header=True).value = df_new

        sheet.range("A1").options(index = True).value = df_new
        
        wb.save()
        return True
    except Exception as err:
        log(f"{current_file_name} {err}")
        return False
    
@mcp.tool()
def append_data_live(file_name: str, data: dict, sheet_name: str = "Sheet1") -> bool:
    """Append data to an Excel file even if it is currently open in Excel.
        The 'data' argument will be converted to a pandas dataframe.
            Required format: {
                        column_name: list[object]
                        }

    Args:
        file_name (str): Name of the output file
        data (dict): User's raw input data
        sheet_name (str, optional): Spreadsheet name. Defaults to "Sheet1".
    Returns:
        bool: Acknowledgement whether the process was successful
    """
    df_new = data_preprocess(file_name, sheet_name, data)

    try:
        wb = xw.Book(file_name)
        sheet = wb.sheets[sheet_name]

        last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end('up').row
        
        # if mode == "write":
        #     sheet.range("A1").options(index = True).value = df_new
        # else:
        #     sheet.range(f"A{last_row + 1}").options(index=True, header=True).value = df_new

        sheet.range(f"A{last_row + 1}").options(index=True, header=True).value = df_new
        
        wb.save()
        return True
    except Exception as err:
        log(f"{current_file_name} {err}")
        return False
    
@mcp.tool()
def update_data_live(file_name: str, data: dict[str, object], index: int, sheet_name: str = "Sheet1") -> bool:
    """updates data in the excel 'table'
       The 'data' argument must contain each column and its corresponding updated/non-updated value
    Args:
        file_name (str): name of the excel file
        data (dict[str, object]): new data
        index (int): row index of the data to be updated
        sheet_name (str, optional): name of the sheet in the file. Defaults to "Sheet1".

    Returns:
        bool: Acknowlegement
    """
    try:
        wb = xw.Book(file_name)
        sheet = wb.sheets[sheet_name]

        df = sheet.range("A1").options(pd.DataFrame, expand = "table", index = False).value
        
        if index < 0 or index >= len(df):
            log(f"{current_file_name} Index {index} out of bounds")
            return False

        row = df.iloc[index]
        updated_row = update_row(row, data)
        
        # DataFrame row index to Excel row:
        # Header is row 1, df.iloc[0] is row 2
        excel_row = index + 2
        
        # Write the updated values back to the specific row
        sheet.range(f"A{excel_row}").value = updated_row.values
        
        wb.save()
    except Exception as err:
        log(f"{current_file_name} {err}")
        return False
    
    return True

@mcp.tool()
def add_line_chart(file_name: str, y: list[float], x: list[float] = None, cell_name: str = "A20", chart_name: str = "Line Chart", sheet_name: str = "Sheet1") -> bool:
    """Adds a line chart to the specified Excel sheet.
    
    Args:
        file_name (str): Name of the Excel file
        y (list[float]): Y-axis data points
        x (list[float], optional): X-axis data points
        cell_name (str, optional): Target cell for the chart. Defaults to "A20".
        chart_name (str, optional): Name of the chart. Defaults to "Line Chart".
        sheet_name (str, optional): Target spreadsheet name. Defaults to "Sheet1".
    """
    try:
        wb = xw.Book(file_name)
        sheet = wb.sheets[sheet_name]
        
        figure = plot_line_chart(y, x)
        
        if figure:
            sheet.pictures.add(figure, name = chart_name, update = True, top = sheet.range(cell_name).top, left = sheet.range(cell_name).left)
            wb.save()
            return True
    except Exception as err:
        log(f"{current_file_name} {err}")
        return False

    return False

@mcp.tool()
def add_pie_chart(file_name: str, labels: list[str], sizes: list[float], cell_name: str = "A20", chart_name: str = "Pie Chart", sheet_name: str = "Sheet1") -> bool:
    """Adds a pie chart to the specified Excel sheet.
    
    Args:
        file_name (str): Name of the Excel file
        labels (list[str]): Labels for each slice
        sizes (list[float]): Sizes for each slice
        cell_name (str, optional): Target cell for the chart. Defaults to "A20".
        chart_name (str, optional): Name of the chart. Defaults to "Pie Chart".
        sheet_name (str, optional): Target spreadsheet name. Defaults to "Sheet1".
    """
    try:
        wb = xw.Book(file_name)
        sheet = wb.sheets[sheet_name]
        
        figure = plot_pie_chart(labels, sizes)
        
        if figure:
            sheet.pictures.add(figure, name = chart_name, update = True, top = sheet.range(cell_name).top, left = sheet.range(cell_name).left)
            wb.save()
            return True
    except Exception as err:
        log(f"{current_file_name} {err}")
        return False

    return False

if __name__ == "__main__":
    mcp.run()
