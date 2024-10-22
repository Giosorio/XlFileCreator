import pandas as pd
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.workbook.protection import WorkbookProtection

from typing import Optional, Union, Callable, Protocol

from .conditional_formatting import highlight_mandatory
from .formats import format_lock_config_dict
from .header_format import set_headers_format


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


def set_formula(df: pd.DataFrame, data_index: int, df_settings: pd.DataFrame) -> pd.DataFrame:
    """
    Set up Excel formulas 

    df: dataframe used to create the template header=None (df_header + df_data_only + df_extra_rows)
    data_index: interger index where the data starts in the df_data
    df_settings: dataframe contaning all the specifications for the template  
    """
    if 'formula' not in df_settings.index:
        return df
    if all(i == '' for i in df_settings.loc['formula']):
        return df

    formula_settings_row = df_settings.loc['formula'].tolist()
    for col, formula_ in zip(df.columns, formula_settings_row):
        if formula_ != '':
            df.iloc[data_index: ,col] = formula_

    return df


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
    data_index: int, df: pd.DataFrame, sheet_password: str) -> None:
    """
    Sets up the format of each column in the dataframe 
    initial_index -> data frame index from which the data starts, EXCLUDING THE HEADER (assuming the header willl be locked)
    """

    # locked = wb.add_format({'locked': True})
    unlocked_text = wb.add_format({'locked': False, 'text_wrap':True})

    for col, header in enumerate(df.columns):
        unlocked_cells = df.iloc[data_index:, col]        
        ws.write_column(data_index, col, unlocked_cells, cell_format=unlocked_text)

    ws.protect(sheet_password)


### VERSION 2
def lock_sheet(wb: xlsxwriter.workbook.Workbook, ws: xlsxwriter.worksheet.Worksheet, 
    data_index: int, df: pd.DataFrame, df_settings: pd.DataFrame, allow_input_extra_rows: bool, 
    sheet_password: str) -> Union[lock_sheet_simple_func, None]:
    """
    If 'lock_sheet_config' is not in the index of the dataframe, all excel columns will be editable 
    If 'lock_sheet_config' contains only blanks, all excel columns will be editable 
    if 'lock_sheet_config' contains only unrecognisable formats, all excel columns will be editable

    If the format is unrecognised or left blank (''), the Excel column will be locked 
    If allow_input_extra_rows=True and the column format is unrecognised or left blank (''), the column 
    will be locked and ONLY the extra rows in the column will be editable

    Comms:
    when concatenating df_data + extra_rows, extra_rows.index starts with 0 to num_extra_rows
    and is stored in the custom index is created from the begining 
    That's why df.loc[0] is referring to the frist blank row added 
    """

    if 'lock_sheet_config' not in df_settings.index:
        return lock_sheet_simple(wb, ws, data_index, df, sheet_password)
    else:
        lock_sheet_config = [config_format if config_format in format_lock_config_dict.keys() else '' for config_format in df_settings.loc['lock_sheet_config']]
        all_blanks = all('' == _format for _format in lock_sheet_config)
        if all_blanks:
            return lock_sheet_simple(wb, ws, data_index, df, sheet_password)

    if allow_input_extra_rows:
        first_blank_row_index = df.index.tolist().index(0)

    for col, lock_config in zip(df.columns, lock_sheet_config):
        if allow_input_extra_rows:
            if lock_config not in format_lock_config_dict.keys():
                ### range from which blank rows start
                unlocked_cells = df.loc[0:, col]
                ws.write_column(first_blank_row_index, col, unlocked_cells, cell_format=wb.add_format(format_lock_config_dict['unlocked_text']))
            else:
                unlocked_cells = df.iloc[data_index:, col]
                ws.write_column(data_index, col, unlocked_cells, cell_format=wb.add_format(format_lock_config_dict[lock_config]))
        else:
            if lock_config in format_lock_config_dict.keys():
                unlocked_cells = df.iloc[data_index:, col]        
                ws.write_column(data_index, col, unlocked_cells, cell_format=wb.add_format(format_lock_config_dict[lock_config]))          

    ws.protect(sheet_password)


class XlFileTemp(Protocol):
    ...


def process_template(writer: pd.ExcelWriter, template: XlFileTemp, split_by_value: bool, template_name: str, 
    split_by: str, split_value: str, sheet_password: Optional[str]=None) -> None:
    """
    Transform the template into the excel file 

    writer: pd.ExcelWriter, Context manager that creates the Excel file
    template: XlFileTemp object
    split_by_value: A boolean flag (True or False). If True, the method filters by the split_value provided. If False, it uses all values from the split_by column.
    template_name: Name of the main sheet of the template in the excel file by default 'Sheet1' -> 'Sheet{j}
    split_by: The name of the column to filter by.
    split_value: The specific value to filter the data by. If set split_value=False it will set the split_value to all records in the split_by column.
    sheet_password: sheet password for the excel file to avoid the users to change the format of the main sheet, default=None 
    """

    df = template.template_filtered(split_by=split_by, split_value=split_value, split_by_value=split_by_value)
    df = set_formula(df, template.data_index, template.df_settings)

    
    df.to_excel(writer, sheet_name=template_name, index=False, header=False)
    if template.dv_config1.df_data_validation is not None: 
        template.dv_config1.df_data_validation.to_excel(writer,sheet_name=template.dv_config1.dropdown_list_sheet, index=False)
        ws_dv = writer.sheets[template.dv_config1.dropdown_list_sheet]
        ws_dv.hide()

    if template.dv_config2.data_validation_dict is not None: 
        template.dv_config2.picklists.to_excel(writer,sheet_name=template.dv_config2.dropdown_list_sheet, index=False)
        ws_dv2 = writer.sheets[template.dv_config2.dropdown_list_sheet]
        ws_dv2.hide()

    wb = writer.book
    ws = writer.sheets[template_name]
    
    ### Insert Header format
    set_headers_format(wb, ws, df, template.df_settings, template.header_index_list, template.hd_index)

    ### Insert Dropdown lists
    template.dv_config1.set_data_validation(ws, df)
    template.dv_config2.set_data_validation(ws, df)

    ### Set Conditional Formatting
    ## The order of the conditions matters. A new condition do not overwrite a previous condition.
    ## The conditions in the conditional_formatting sheet are superimposed over the Mandatory fields
    ## The mandtory flag does not overwrite an existing condition in the conditional_formatting sheet
    template.cond_formatting.set_conditional_formatting(wb, ws, df)
    highlight_mandatory(wb, ws, df, template.df_settings, template.data_index, template.extra_rows, template.num_rows_extra)

    ### Set column width
    column_width(ws, df, template.df_settings)

    ### Protect Sheet
    ### All sheets will have the password
    if sheet_password is not None and sheet_password != '':
        ### Hide all rows without data. Even when the empty extra rows are allowed
        ## it will only show those that can be filled in
        ws.set_default_row(hide_unused_rows=True)
        
        ### Hide unused columns 
        last_col_num = df.columns[-1]
        hide_from_col_name = xlsxwriter.utility.xl_col_to_name(last_col_num + 1)
        ws.set_column(f'{hide_from_col_name}:XFD', None, None, {"hidden": True})

        lock_sheet(wb, ws, template.data_index, df, template.df_settings, template.extra_rows, sheet_password)


def create_xl_file(*, template: XlFileTemp, file_path: str, template_name: str, split_by_value: Optional[bool]=None, split_by: Optional[str]=None,
    split_value: Optional[str]=None, sheet_password: Optional[str]=None, workbook_password: Optional[str]=None) -> None:
    """
    Creates the context manager pd.ExcelWriter (writer) to create the excel file of the template (XlFileTemp).

    template: XlFileTemp object
    file_path: complete filename of the excel file
    template_name: Name of the main sheet of the template in the excel file by default 'Sheet1' -> 'Sheet{j}
    split_by: The name of the column to filter by.
    split_value: The specific value to filter the data by. If set split_value=False it will set the split_value to all records in the split_by column.
    split_by_value: A boolean flag (True or False). If True, the method filters by the split_value provided. If False, it uses all values from the split_by column.
    sheet_password: sheet password for the excel file to avoid the users to change the format of the main sheet, default=None 
    workbook_password: workbook password to avoid the users to add more sheets in the excel file, defaul=None
    """
    
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        process_template(writer, template, split_by_value, template_name, split_by, split_value, sheet_password)
        
     
    ### Protect Workbook
    if workbook_password is not None and workbook_password != '':
        protect_workbook(file_path, password=workbook_password)


