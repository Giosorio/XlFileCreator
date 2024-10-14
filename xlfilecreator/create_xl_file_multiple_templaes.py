import pandas as pd
import xlsxwriter
from tqdm.auto import tqdm 

import datetime
import os
import shutil
from typing import Optional, List

from .conditional_formatting import highlight_mandatory
from .create_xlfile import (set_formula, column_width, lock_sheet, protect_workbook)
from .encrypt_xl import set_password, create_password
from .header_format import set_headers_format
from .utils_func import (set_project_name, create_output_folders)
from .xlfiletemp import XlFileTemp


def create_xl_file_multiple_temp(*, project_name: str, template_list: List[XlFileTemp], split_by: Optional[str]=None, split_by_range: Optional[List[str]]=None, 
    batch: Optional[int]=1, sheet_password: Optional[str]=None, workbook_password: Optional[str]=None,
    protect_files: Optional[bool]=False, random_password: Optional[bool]=False, in_zip: Optional[bool]=False) -> None:
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

    if split_by is None and split_by_range is None:
        return None

    if isinstance(split_by_range, list):
        values_to_split = set(split_by_range)
    else:
        raise TypeError(f'{split_by_range} is not a list')

    ### Check feasibility
    for template in template_list:
        print(f"Checking: {template.tab_names['main_sheet']}")
        template.check_split_by_range(split_by, split_by_range)

    ### Create output folders
    today = datetime.datetime.today().strftime('%Y%m%d')
    project = set_project_name(project_name)
    path_1, path_2 = create_output_folders(project.name, today, protect_files)
    
    ### 
    password_master = []
    pbar = tqdm(total=len(values_to_split))
    for i, split_value in enumerate(values_to_split, 1):
        ### Remove special characters from the supplier name
        name = ''.join(char for char in split_value if char == ' ' or char.isalnum())
        id_file = f'{project.name}ID{batch}{i:03d}'
        file_name = f'{id_file}-{name}-{today}.xlsx'
        file_path = f'{path_1}/{file_name}'
        
        for j, template in enumerate(template_list, 1):
            df = template.template_filtered(split_by=split_by, split_value=split_value)
            df = set_formula(df, template.df_settings, template.extra_rows)

            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name=f'Sheet{j}', index=False, header=False)
                if template.dv_config1.df_data_validation is not None: 
                    template.dv_config1.df_data_validation.to_excel(writer,sheet_name=template.dv_config1.dropdown_list_sheet, index=False)
                    ws_dv = writer.sheets[template.dv_config1.dropdown_list_sheet]
                    ws_dv.hide()

                if template.dv_config2.data_validation_dict is not None: 
                    template.dv_config2.picklists.to_excel(writer,sheet_name=template.dv_config2.dropdown_list_sheet, index=False)
                    ws_dv2 = writer.sheets[template.dv_config2.dropdown_list_sheet]
                    ws_dv2.hide()

                wb = writer.book
                ws = writer.sheets[f'Sheet{j}']
                
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
                
        ### Protect Workbook
        if workbook_password is not None and workbook_password != '':
            protect_workbook(file_path, password=workbook_password)

        ### Create Password master df
        if protect_files is True:
            pw = create_password(project, split_value, random_password)    
            password_master.append((id_file, file_name, split_value, pw))

    ### Encrypt Excel files
    if protect_files is True:
        df_pw = pd.DataFrame(password_master, columns=['File ID', 'Filename', split_by, 'Password'])
        passwordMaster_name = f'{project.name}-PasswordMaster-{today}.csv'
        df_pw.to_csv(passwordMaster_name, index=False)

        set_password(path_1, path_2, passwordMaster_name)
        print(df_pw)

    pbar.close()

    if in_zip:
        shutil.make_archive(path_1, 'zip', path_1)
        shutil.make_archive(path_2, 'zip', path_2)
        os.system(f'rm -r {path_1}')
        os.system(f'rm -r {path_2}')
