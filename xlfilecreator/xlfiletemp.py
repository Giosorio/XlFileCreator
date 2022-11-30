import pandas as pd

import datetime
import os
import shutil
from typing import List, Dict, Optional

from .create_xlfile import create_xl_file
from .config_file import config_file
from .data_validation import clean_df_data_validation, get_data_validation_dict
from .encrypt_xl import set_password, create_password
from .header_format import set_headers_format
from .utils_func import get_google_sheet_df, get_headers, get_df_data, rows_extra, set_project_name
from .utils_func import create_output_folders, clean_df_main, get_google_sheet_validation, get_column_to_split_by


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

    def __init__(self, df_main: pd.DataFrame, df_data_validation_complete: Optional[pd.DataFrame]=None, 
    allow_input_extra_rows: Optional[bool]=False, dropdown_list_sheet: Optional[str]='Dropdown_Lists') -> None:

        self.__df_data = None
        self.df_data_only = df_main[df_main.index=='']
        self.df_settings = df_main[df_main.index!='']
        self.extra_rows = allow_input_extra_rows
        self.__last_extra_rows = allow_input_extra_rows
        self.header_index_list, self.df_hd = get_headers(self.df_settings)
        self.hd_index = self.df_data.index.tolist().index('HEADER')
        self.data_index = self.df_data.index.tolist().index('')
        
        self.dropdown_list_sheet = dropdown_list_sheet
        self.df_data_validation_complete, self.df_data_validation = clean_df_data_validation(df_data_validation_complete, self.df_settings)
        if df_data_validation_complete is None:
            self.data_validation_dict = None
            self.data_val_headers = None
        else:
            self.data_validation_dict, self.data_val_headers = get_data_validation_dict(self.df_settings, self.df_data_validation_complete, self.df_data_validation, self.dropdown_list_sheet)
        

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
    def load_from_excel(cls, xl_file):
        """Constructor of XlFileTemp
        Creates an XlFileTemp object from an excel file"""
        pass
    
    @classmethod
    def read_google_sheets_file(cls, sheet_id: str, sheet_name: str, dropdown_list_sheet: Optional[str]=None):
        """
        Returns a XlFileTemp object

        Parameters
        sheet_id: google sheets id 
        sheet_name: name of the sheet where the data is stored
        dropdown_list_sheet: name of the sheet where the dropdownlists and data validation settings are located, default=None
        """

        ### Read google sheets file
        df_main = get_google_sheet_df(sheet_id, sheet_name)
        df_main = clean_df_main(df_main)
        if dropdown_list_sheet is None or dropdown_list_sheet == '':
            df_data_validation_complete = None
        else:
            df_data_validation_complete = get_google_sheet_validation(sheet_id, dropdown_list_sheet)

        return cls(df_main, df_data_validation_complete, dropdown_list_sheet=dropdown_list_sheet)

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
            create_xl_file(project_name, self.df_data, self.df_settings, self.data_validation_dict, 
            self.data_val_headers, self.df_data_validation, self.hd_index, self.header_index_list, 
            self.extra_rows, self.dropdown_list_sheet, sheet_password, workbook_password)

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
            create_xl_file(file_path, df_split_value, self.df_settings, self.data_validation_dict, 
            self.data_val_headers, self.df_data_validation, self.hd_index, self.header_index_list, 
            self.extra_rows, self.dropdown_list_sheet, sheet_password, workbook_password)


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
