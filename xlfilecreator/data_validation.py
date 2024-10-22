import pandas as pd
import xlsxwriter

from abc import ABC, abstractclassmethod
from typing import Dict, Union

from .data_validation_config1_func import get_data_validation_dict,clean_df_data_validation
from .data_validation_typing import DataValDict



class DataValidationConfiguration(ABC):

    @abstractclassmethod
    def __init__(self) -> None:
        self.data_validation_dict: dict
        self.data_val_headers: list 
        self.data_index: int

    def set_data_validation(self, ws: xlsxwriter.worksheet.Worksheet, df: pd.DataFrame) -> None:
        
        """
        Set up data validation, dropdown lists 
        Parameters:
        ws: worksheet
        df: dataframe used to create the template header=None
        self.data_validation_dict: DataValDict Dictionary containing the opctions_dict for each field in scope for data validation
        self.data_val_headers: List[Header] List of headers in scope for data validation
        """
        if self.data_validation_dict is None:
            return None

        column_indexes_to_apply_data_validation = [i for i, hd in enumerate(df.loc['HEADER']) if hd in self.data_val_headers]  
        last_row_index = df.shape[0] - 1  

        for col in column_indexes_to_apply_data_validation:
            hd = df.loc['HEADER', col]
            opts_dict = self.data_validation_dict[hd]
            ### ws.data_validation(first_row, first_col, last_row, last_col, options_dict={...})
            # ws.data_validation(initial_index, col, last_row_index, col, {'validate':'list', 'source':data_source_dict[hd], 'error_type':'stop'})
            ws.data_validation(self.data_index, col, last_row_index, col, opts_dict)


class DataValidationConfig1(DataValidationConfiguration):
    """
    df_data_validation_complete: dataframe containing all the dropdown lists and the settings for the data validation
    df_data_validation: dataframe contaning only the dropdown lists 
    data_validation_dict: DataValDict

    """

    def __init__(self, data_index: int, df_dvconfig1: Union[pd.DataFrame,None], dropdown_list_sheet: str, df_settings: pd.DataFrame) -> None:
        if df_dvconfig1 is None:
            self.data_validation_dict = None
            self.data_val_headers = None
            self.df_data_validation_complete = None
            self.df_data_validation = None
            self.data_index: int = None
        else:
            self.data_index: int = data_index
            self.dropdown_list_sheet = dropdown_list_sheet
            self.df_data_validation_complete, self.df_data_validation = clean_df_data_validation(df_dvconfig1, df_settings)
            self.data_val_headers = self.df_data_validation.columns.tolist()
            self.data_validation_dict = get_data_validation_dict(df_settings, self.df_data_validation_complete, self.df_data_validation, self.dropdown_list_sheet)


class DataValidationConfig2(DataValidationConfiguration):

    def __init__(self, data_index: int, df_picklists: Union[pd.DataFrame,None], dropdown_list_sheet: str, df_dvconfig2: Union[pd.DataFrame,None]) -> None:
        if df_dvconfig2 is None or df_picklists is None:
            self.__data_validation_dict = None
            self.data_val_headers = None
            self.picklists = None
            self.data_index: int = None
            self.df_dvconfig2 : pd.DataFrame = None
        else:
            self.data_index: int = data_index
            self.df_dvconfig2 : pd.DataFrame = df_dvconfig2
            self.picklists = df_picklists
            self.dropdown_list_sheet = dropdown_list_sheet
            self.__data_validation_dict = None
            self.data_val_headers = self.data_validation_dict.keys()

    @staticmethod
    def create_opts_dict(opts_settings:pd.Series) -> Dict[str,str]:
        try:
            opts_dict1={'validate': opts_settings['validate'],
                    'source': opts_settings['source'],
                    'error_type': opts_settings['error_type'],
                    'input_title': opts_settings['input_title'],
                    'input_message': opts_settings['input_message'],
                    'error_title': opts_settings['error_title'],
                    'error_message': opts_settings['error_message']}
        except KeyError as ke:
            print('Header not recognised')
            raise ke
        
        opts_dict = {opt:value for opt, value in opts_dict1.items() if value != ''}

        return opts_dict 

    @property
    def data_validation_dict(self) -> DataValDict:

        if self.df_dvconfig2 is None:
            return None

        if self.__data_validation_dict is not None:
            return self.__data_validation_dict

        data_validation_opts_dict = {}
        for row in self.df_dvconfig2.index:
            hd_to_apply = self.df_dvconfig2.loc[row, 'apply_to']
            opts_dict = DataValidationConfig2.create_opts_dict(self.df_dvconfig2.loc[row])
            data_validation_opts_dict[hd_to_apply] = opts_dict
        
        self.__data_validation_dict = data_validation_opts_dict
        return self.__data_validation_dict



        
