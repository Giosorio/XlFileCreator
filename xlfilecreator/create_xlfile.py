import pandas as pd
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.workbook.protection import WorkbookProtection

from typing import List, Dict, Optional, Union, Callable

from .conditional_formatting import highlight_mandatory
from .formats import format_lock_config_dict
from .header_format import set_headers_format
from .data_validation_typing import DataValDict, Header


def protect_workbook(path: str, password: str) -> None:
    """
    Openpyxl -> Manipulate a file that is already created

    PARAMETERS
    path -> Location where the excel file is stored
    password -> workbook password
    """
    
    ### PROTECT WORKBOOK openpyxl
    wb = load_workbook(path)
    wb.security = WorkbookProtection(workbookPassword=password, lockStructure=True)
    wb.save(path)


def column_width(ws: xlsxwriter.worksheet.Worksheet, df: pd.DataFrame, df_settings: pd.DataFrame) -> None:
    """
    Set up the columns width in character units.

    worksheet.set_column(first_col, last_col, width, cell_format, options)

    Parameters:
    ws: worksheet
    df: dataframe used to create the template header=None
    df_settings: dataframe contaning all the specifications for the template  
    """

    width_format = df_settings.loc['column_width'].tolist()
    for col_num, width in zip(df.columns, width_format):
        if width == '':
            width = 25
        ws.set_column(col_num, col_num, width=int(width))


### VERSION 1
lock_sheet_simple_func = Callable[[xlsxwriter.workbook.Workbook, xlsxwriter.worksheet.Worksheet, pd.DataFrame, str], None]
def lock_sheet_simple(wb: xlsxwriter.workbook.Workbook, ws: xlsxwriter.worksheet.Worksheet, 
df: pd.DataFrame, sheet_password: str) -> None:
    """
    Sets up the format of each column in the dataframe 
    initial_index -> data frame index from which the data starts, EXCLUDING THE HEADER (assuming the header willl be locked)
    """

    # locked = wb.add_format({'locked': True})
    unlocked_text = wb.add_format({'locked': False, 'text_wrap':True})

    ### Index where the first blank ('') is loacated in the df.index, that is where the data starts
    initial_index = df.index.tolist().index('')
    for col, header in enumerate(df.columns):
        unlocked_cells = df.iloc[initial_index:, col]        
        ws.write_column(initial_index, col, unlocked_cells, cell_format=unlocked_text)

    ws.protect(sheet_password)


### VERSION 2
def lock_sheet(wb: xlsxwriter.workbook.Workbook, ws: xlsxwriter.worksheet.Worksheet, 
df: pd.DataFrame, df_settings: pd.DataFrame, allow_input_extra_rows: bool, 
sheet_password: str) -> Union[lock_sheet_simple_func, None]:
    """
    If 'lock_sheet_config' is not in the index of the dataframe, all excel columns will be editable 
    If 'lock_sheet_config' contains only blanks, all excel columns will be editable 
    if 'lock_sheet_config' contains only unrecognisable formats, all excel columns will be editable

    If the format is not recognised the excel column will be locked 
    If allow_input_extra_rows=True but the column should be locked, ONLY the extra rows in the column will be editable 

    Comms:
    when concatenating df_data + extra_rows, extra_rows.index starts with 0 to 100 
    and is stored in the custom index is created from the begining 
    That's why df.loc[0] is referring to the frist blank row added 
    """

    if 'lock_sheet_config' not in df_settings.index:
        return lock_sheet_simple(wb, ws, df, sheet_password)
    else:
        lock_sheet_config = [config_format if config_format in format_lock_config_dict.keys() else '' for config_format in df_settings.loc['lock_sheet_config']]
        all_blanks = all('' == _format for _format in lock_sheet_config)
        if all_blanks:
            return lock_sheet_simple(wb, ws, df, sheet_password)


    initial_index = df.index.tolist().index('')
    if allow_input_extra_rows:
        first_blank_row_index = df.index.tolist().index(0)

    for col, lock_config in zip(df.columns, lock_sheet_config):
        if allow_input_extra_rows:
            if lock_config not in format_lock_config_dict.keys():
                unlocked_cells = df.loc[0:, col]                           #### range from which blank rows start
                # unlocked_cells = df.iloc[first_blank_row_index:, col]    #### range from which blank rows start
                ws.write_column(first_blank_row_index, col, unlocked_cells, cell_format=eval(format_lock_config_dict['unlocked_text']))
            else:
                unlocked_cells = df.iloc[initial_index:, col]
                ws.write_column(initial_index, col, unlocked_cells, cell_format=eval(format_lock_config_dict[lock_config]))
        else:
            if lock_config in format_lock_config_dict.keys():
                unlocked_cells = df.iloc[initial_index:, col]        
                ws.write_column(initial_index, col, unlocked_cells, cell_format=eval(format_lock_config_dict[lock_config]))          

    ws.protect(sheet_password)


def set_data_validation(ws: xlsxwriter.worksheet.Worksheet, df: pd.DataFrame, 
data_validation_opts_dict: DataValDict, data_val_headers: List[Header]) -> None:
    """
    Set up data validation, dropdown lists 
    Parameters:
    ws: worksheet
    df: dataframe used to create the template header=None
    data_validation_opts_dict: Dictionary containing the opctions_dict for each field in scope for data validation
    data_val_headers: dataframe containing only the dropdownlists 
    """

    column_indexes_to_apply_data_validation = [i for i, hd in enumerate(df.loc['HEADER']) if hd in data_val_headers]  
    initial_index = df.index.tolist().index('')  ## df index 0 = excel row 1
    last_row_index = df.shape[0] - 1  

    for col in column_indexes_to_apply_data_validation:
        hd = df.loc['HEADER', col]
        opts_dict = data_validation_opts_dict[hd]
        ### ws.data_validation(first_row, first_col, last_row, last_col, options_dict={...})
        # ws.data_validation(initial_index, col, last_row_index, col, {'validate':'list', 'source':data_source_dict[hd], 'error_type':'stop'})
        ws.data_validation(initial_index, col, last_row_index, col, opts_dict)


def create_xl_file(file_path: str, df: pd.DataFrame, df_settings: pd.DataFrame, data_validation_opts_dict: DataValDict, 
data_val_headers: List[str], df_data_validation: pd.DataFrame, header_index: int, data_index: int, header_index_list: List[str], allow_input_extra_rows: Optional[bool]=False, 
dropdown_list_sheet: Optional[str]=None, sheet_password: Optional[str]=None, workbook_password: Optional[str]=None) -> None:
    """
    file_path: complete filename of the excel file
    df: dataframe containing only the headers and data of the main sheet of the excel file
    df_settings: dataframe containing the config requirements for the main sheet (width, header_format, description_header...)
    data_validation_opts_dict: dictionary where the keys are the headers to apply data validation and the values are the dictionaries containing the options for the data validation 
    data_val_headers: list of headers/columns to apply data validation
    df_data_validation: dataframe containing the dropdown lists for data validation ready to be set as a second sheet in the excel file 
    header_index_list: list of headers included in the index ['Description_header', 'HEADER', 'Example_header']
    dropdown_list_sheet: Name of the sheet contaning the dropdown lists and settings for data validation, default=None 
    sheet_password: sheet password for the excel file to avoid the users to change the format of the main sheet, default=None 
    workbook_password: workbook password to avoid the users to add more sheets in the excel file, defaul=None
    """
    
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
        if df_data_validation is not None: 
            df_data_validation.to_excel(writer,sheet_name=dropdown_list_sheet, index=False)
            ws_dv = writer.sheets[dropdown_list_sheet]
            ws_dv.hide()

        wb = writer.book
        ws = writer.sheets['Sheet1']
        
        ### Insert Header format
        set_headers_format(wb, ws, df, df_settings, header_index_list, header_index)

        ### Insert Dropdown lists
        if df_data_validation is not None: 
            set_data_validation(ws, df, data_validation_opts_dict, data_val_headers)

        ### Set Conditional Formatting
        highlight_mandatory(wb, ws, df, df_settings, data_index, allow_input_extra_rows)

        ### Set column width
        column_width(ws, df, df_settings)

        ### Protect Sheet
        if sheet_password is not None and sheet_password != '':
            lock_sheet(wb, ws, df, df_settings, allow_input_extra_rows, sheet_password)

    ### Protect Workbook
    if workbook_password is not None and workbook_password != '':
        protect_workbook(file_path, password=workbook_password)




