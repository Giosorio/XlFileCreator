import pandas as pd

import datetime
import os
import shutil
from typing import Optional

from .create_xlfile import create_xl_file
from .conditional_formatting import CondFormatting
from .config_file import config_file
from .data_validation import DataValidationConfig1, DataValidationConfig2
from .encrypt_xl import set_password, create_password
from .utils_func import (get_google_sheet_df, get_headers, get_df_data, check_google_sh_reader,rows_extra,
                        set_project_name, get_google_sheet_validation2, get_excel_dvalidation2,
                        create_output_folders, clean_df_main, get_google_sheet_validation, 
                        get_column_to_split_by, get_excel_df)


class XlFileTemp:
    """
    Properties:

    df_hd: dataframe contaning the headers (description_header, example_header, header)
    df_data: dataframe containing the headers + data + extra_rows(optional)
    df_settings: dataframe containing configuration of the excel file (format, width, lock columns, etc)
    header_index_list: list of header indexes in scope i.e ['Description_header','HEADER','example_row'] 
    hd_index: interger index where the header is located in the df_data
    data_index: interger index where the data starts in the df_data
    lenght: number of rows of the data 
    df_data_validation_complete (optional): dataframe containing all the dropdown lists and the settings for the data validation
    df_data_validation (optional): dataframe contaning only the dropdown lists 
    
    Methods:

    read_google_sheets_file(cls): Creates a XlFileTemp object from a google sheeets workbook
    """

    def __init__(self, df_main: pd.DataFrame, df_dvconfig1: Optional[pd.DataFrame]=None, df_dvconfig2: Optional[pd.DataFrame]=None,
    allow_input_extra_rows: Optional[bool]=False, data_validation_sheet_config1: Optional[str]='Dropdown_Lists',
    dropdown_lists_sheet_config2: Optional[str]='Dropdown_Lists_2', df_picklists: Optional[pd.DataFrame]=None,
    df_condf: Optional[pd.DataFrame]=None) -> None:

        self.__df_data = None
        self.df_data_only = df_main[df_main.index=='']
        self.df_settings = df_main[df_main.index!='']
        self.extra_rows = allow_input_extra_rows
        self.__last_extra_rows = allow_input_extra_rows
        self.header_index_list, self.df_hd = get_headers(self.df_settings)
        self.hd_index = self.df_data.index.tolist().index('HEADER')
        self.data_index = self.df_data.index.tolist().index('')
        
        self.data_validation_sheet_config1 = data_validation_sheet_config1
        self.dv_config1 = DataValidationConfig1(df_dvconfig1, data_validation_sheet_config1, self.df_settings)
        
        self.dropdown_lists_sheet_config2 = dropdown_lists_sheet_config2
        self.dv_config2 = DataValidationConfig2(df_picklists, dropdown_lists_sheet_config2, df_dvconfig2)

        self.cond_formatting = CondFormatting(df_condf, self.df_data)

    @property
    def df_data(self) -> pd.DataFrame:
        if self.__df_data is None or self.extra_rows != self.__last_extra_rows:
            self.__df_data = get_df_data(self.df_hd, self.df_data_only, allow_input_extra_rows=self.extra_rows)
            if self.extra_rows != self.__last_extra_rows:
                self.__last_extra_rows = self.extra_rows 
                print(f'Update: allow_input_extra_rows= {self.extra_rows}')

        return self.__df_data

    @property
    def length(self) -> int:
        """lenght: number of rows of the data """
        
        return self.df_data.shape[0]
    
    @classmethod
    def read_excel(cls, xl_file: str, sheet_name: str, data_validation_sheet_config1: Optional[str]=None,
        data_validation_sheet_config2: Optional[str]=None, dropdown_lists_sheet_config2: Optional[str]=None,
        conditional_formatting_sheet: Optional[str]=None):
        """
        Constructor of XlFileTemp
        Creates an XlFileTemp object from an excel file

        Parameters
        xl_file: Excel file path
        sheet_name: name of the sheet where the data is stored
        dropdown_list_sheet: name of the sheet where the dropdownlists and data validation settings are located, default=None
        """
        
        df_main = get_excel_df(xl_file, sheet_name)
        df_main = clean_df_main(df_main)
        if conditional_formatting_sheet is None or conditional_formatting_sheet == '':
            df_condf = None
        else:
            df_condf = pd.read_excel(xl_file, sheet_name=conditional_formatting_sheet, na_filter=False)

        if data_validation_sheet_config1 is None or data_validation_sheet_config1 == '':
            df_dvconfig1 = None
        else:
            df_dvconfig1 = get_excel_df(xl_file, sheet_name=data_validation_sheet_config1, header='HEADER')
        
        df_dvconfig2, df_picklists = get_excel_dvalidation2(xl_file, data_validation_sheet_config2, dropdown_lists_sheet_config2)
        
        return cls(df_main, df_dvconfig1, df_dvconfig2, data_validation_sheet_config1=data_validation_sheet_config1, 
                dropdown_lists_sheet_config2=dropdown_lists_sheet_config2, df_picklists=df_picklists, df_condf=df_condf)

    @classmethod
    def read_google_sheets_file(cls, sheet_id: str, sheet_name: str, data_validation_sheet_config1: Optional[str]=None,
        data_validation_sheet_config2: Optional[str]=None, dropdown_lists_sheet_config2: Optional[str]=None,
        conditional_formatting_sheet: Optional[str]=None):
        """
        Returns a XlFileTemp object

        Parameters
        sheet_id: google sheets id 
        sheet_name: name of the sheet where the data is stored
        df_dvconfig1: dataframe Data Validation Config 1
        dropdown_list_sheet: name of the sheet where the dropdownlists and data validation settings are located, default=None
        """

        ### Read google sheets file
        df_main = get_google_sheet_df(sheet_id, sheet_name)
        df_main = clean_df_main(df_main)
        df_dvconfig1 = get_google_sheet_validation(sheet_id, data_validation_sheet_config1)
        df_dvconfig2, df_picklists = get_google_sheet_validation2(sheet_id, data_validation_sheet_config2, dropdown_lists_sheet_config2)
        df_condf = check_google_sh_reader(sheet_id, conditional_formatting_sheet, na_filter=False, header=0, index_col=None)

        return cls(df_main, df_dvconfig1, df_dvconfig2, data_validation_sheet_config1=data_validation_sheet_config1, 
                dropdown_lists_sheet_config2=dropdown_lists_sheet_config2, df_picklists=df_picklists, df_condf=df_condf)

    @staticmethod
    def export_config_file() -> None:
        """
        Export Excel file used as a tempalte to create an XlFileTemp object
        The template includes data that can be pass as the parameter of the 
        constructor 'read_google_sheets_file()' to create an XlFileTemp object
        """
        config_file()

    def to_excel(self, project_name: Optional[str]=None, split_by: Optional[str]=None, batch: Optional[int]=1, 
        sheet_password: Optional[str]=None, workbook_password: Optional[str]=None, allow_input_extra_rows: Optional[bool]=None, 
        protect_files: Optional[bool]=False, random_password: Optional[bool]=False, in_zip: Optional[bool]=False) -> None:
        """
        Creates the excel file
        project_name: name of the project, it will be part of the filename of the templates. If split_by is None it will be the name of the single file generated
        split_by: Name of the column to filter and create new templates
        batch: Number of the batch. Included in the filename of the templates 
        sheet_password: sheet password for the excel file to avoid the users to change the format of the main sheet, default=None 
        workbook_password: workbook password to avoid the users to add more sheets in the excel file, defaul=None
        allow_input_extra_rows: False/True Determines if the templates allow the user to fill out more rows in the template
        protect_files: False/True encrypt the files
        random_password: False/True if protect_files is True it determines if the password of the files should be random or based on a logic
        in_zip: False/True Download folders in zip 
        """

        today = datetime.datetime.today().strftime('%Y%m%d')

        if allow_input_extra_rows is not None:
            self.extra_rows = allow_input_extra_rows 
        
        if project_name is None or '':
            project_name = f'Project-{today}'

        if split_by is None:
            if not project_name.endswith('.xlsx'):
                project_name = project_name + '.xlsx'

            create_xl_file(project_name, self.df_data, self.df_settings, self.dv_config1, self.dv_config2,
            self.cond_formatting, self.hd_index, self.data_index, self.header_index_list, self.extra_rows, 
            sheet_password, workbook_password)
            return None

        if self.extra_rows or allow_input_extra_rows:
            df_rows_extra = rows_extra(self.df_data_only)
        else:
            df_rows_extra = None

        project = set_project_name(project_name)
        path_1, path_2 = create_output_folders(project.name, today, protect_files)

        ### Unique list of values to split 
        col_to_split = get_column_to_split_by(self.df_settings, split_by)
        values_to_split = set(self.df_data_only[col_to_split])
        print('Number of files: ', len(values_to_split))

        password_master = []
        for i, split_value in enumerate(values_to_split,1):
            ### Filter the values to include in the template
            df_split_value = self.df_data[self.df_data[col_to_split]==split_value]
            
            ### Include the headers on the top
            df_split_value = pd.concat([self.df_hd, df_split_value, df_rows_extra])

            ### Remove special characters from the supplier name
            name = ''.join(char for char in split_value if char == ' ' or char.isalnum())
            id_file = f'{project.name}ID{batch}{i:03d}'
            file_name = f'{id_file}-{name}-{today}.xlsx'
            file_path = f'{path_1}/{file_name}'

            ### Create Excel file
            create_xl_file(file_path, df_split_value, self.df_settings, self.dv_config1, self.dv_config2, 
            self.cond_formatting, self.hd_index, self.data_index, self.header_index_list, self.extra_rows, 
            sheet_password, workbook_password)

            ### Create Password master df
            if protect_files is True:
                pw = create_password(project, split_value, random_password)    
                password_master.append((id_file, file_name, split_value, pw))

        ### Encrypt Excel files
        if protect_files is True:
            df_pw = pd.DataFrame(password_master, columns=['File ID', 'Filename', 'Supplier', 'Password'])
            passwordMaster_name = f'{project.name}-PasswordMaster-{today}.csv'
            df_pw.to_csv(passwordMaster_name, index=False)

            set_password(path_1, path_2, passwordMaster_name)
        
        if in_zip:
            shutil.make_archive(path_1, 'zip', path_1)
            shutil.make_archive(path_2, 'zip', path_2)
            os.system(f'rm -r {path_1}')
            os.system(f'rm -r {path_2}')
