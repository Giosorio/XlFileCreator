import pandas as pd
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.workbook.protection import WorkbookProtection

from typing import List, Optional, Union, Callable, Protocol

from .conditional_formatting import highlight_mandatory, CondFormatting
from .formats import format_lock_config_dict
from .header_format import set_headers_format
from .data_validation import DataValidationConfig1, DataValidationConfig2


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


def set_formula(df: pd.DataFrame, df_settings: pd.DataFrame, allow_input_extra_rows: Optional[bool]=False) -> pd.DataFrame:
    
    if 'formula' not in df_settings.index:
        return df
    if all(i == '' for i in df_settings.loc['formula']):
        return df

    if allow_input_extra_rows:
        ### Identify where the extra rows start
        first_blank_row = df.index.tolist().index(0)
        ### Remove index from the extra rows [0,1,2...] -> ''
        df.reset_index(inplace=True)
        df['index'].iloc[first_blank_row:] = ''
        df.set_index('index', inplace=True)

    df_hd = df[df.index != ''].copy(deep=True)
    df_data = df[df.index == ''].copy(deep=True)

    formula_settings_row = df_settings.loc['formula'].tolist()
    for col, formula_ in zip(df.columns, formula_settings_row):
        if formula_ != '':
            df_data[col] = formula_

    if allow_input_extra_rows:
        ### Mark with 0 the index where the extra rows start 
        df_data.reset_index(inplace=True)
        df_data['index'].iloc[first_blank_row] = 0
        df_data.set_index('index', inplace=True)


    return pd.concat([df_hd, df_data])


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
                ws.write_column(first_blank_row_index, col, unlocked_cells, cell_format=wb.add_format(format_lock_config_dict['unlocked_text']))
            else:
                unlocked_cells = df.iloc[initial_index:, col]
                ws.write_column(initial_index, col, unlocked_cells, cell_format=wb.add_format(format_lock_config_dict[lock_config]))
        else:
            if lock_config in format_lock_config_dict.keys():
                unlocked_cells = df.iloc[initial_index:, col]        
                ws.write_column(initial_index, col, unlocked_cells, cell_format=wb.add_format(format_lock_config_dict[lock_config]))          

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
    df = set_formula(df, template.df_settings, template.extra_rows)

    
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

        lock_sheet(wb, ws, df, template.df_settings, template.extra_rows, sheet_password)


def create_xl_file(file_path: str, df: pd.DataFrame, df_settings: pd.DataFrame, dv_config1: DataValidationConfig1, dv_config2: DataValidationConfig2,
cond_formatting: CondFormatting, header_index: int, data_index: int, header_index_list: List[str], allow_input_extra_rows: Optional[bool]=False, 
num_rows_extra: Optional[int]=100, sheet_password: Optional[str]=None, workbook_password: Optional[str]=None) -> None:
    """
    file_path: complete filename of the excel file
    df: dataframe containing only the headers and data of the main sheet of the excel file
    df_settings: dataframe containing the config requirements for the main sheet (width, header_format, description_header...)
    dv_config1: DataValidationConfig1 object containing the configuration for Data Validation 1
    dv_config2: DataValidationConfig2 object containing the configuration for Data Validation 2
    header_index_list: list of headers included in the index ['Description_header', 'HEADER', 'Example_header']
    sheet_password: sheet password for the excel file to avoid the users to change the format of the main sheet, default=None 
    workbook_password: workbook password to avoid the users to add more sheets in the excel file, defaul=None
    """

    df = set_formula(df, df_settings, allow_input_extra_rows)
    
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
        if dv_config1.df_data_validation is not None: 
            dv_config1.df_data_validation.to_excel(writer,sheet_name=dv_config1.dropdown_list_sheet, index=False)
            ws_dv = writer.sheets[dv_config1.dropdown_list_sheet]
            ws_dv.hide()

        if dv_config2.data_validation_dict is not None: 
            dv_config2.picklists.to_excel(writer,sheet_name=dv_config2.dropdown_list_sheet, index=False)
            ws_dv2 = writer.sheets[dv_config2.dropdown_list_sheet]
            ws_dv2.hide()

        wb = writer.book
        ws = writer.sheets['Sheet1']
        
        ### Insert Header format
        set_headers_format(wb, ws, df, df_settings, header_index_list, header_index)

        ### Insert Dropdown lists
        dv_config1.set_data_validation(ws, df)
        dv_config2.set_data_validation(ws, df)

        ### Set Conditional Formatting
        ## The order of the conditions matters. A new condition do not overwrite a previous condition.
        ## The conditions in the conditional_formatting sheet are superimposed over the Mandatory fields
        ## The mandtory flag does not overwrite an existing condition in the conditional_formatting sheet
        cond_formatting.set_conditional_formatting(wb, ws, df)
        highlight_mandatory(wb, ws, df, df_settings, data_index, allow_input_extra_rows, num_rows_extra)

        ### Set column width
        column_width(ws, df, df_settings)

        ### Protect Sheet
        if sheet_password is not None and sheet_password != '':
            ### Hide all rows without data. Even when the empty extra rows are allowed
            ## it will only show those that can be filled in
            ws.set_default_row(hide_unused_rows=True)
            
            ### Hide unused columns 
            last_col_num = df.columns[-1]
            hide_from_col_name = xlsxwriter.utility.xl_col_to_name(last_col_num + 1)
            ws.set_column(f'{hide_from_col_name}:XFD', None, None, {"hidden": True})

            lock_sheet(wb, ws, df, df_settings, allow_input_extra_rows, sheet_password)
     
    ### Protect Workbook
    if workbook_password is not None and workbook_password != '':
        protect_workbook(file_path, password=workbook_password)




