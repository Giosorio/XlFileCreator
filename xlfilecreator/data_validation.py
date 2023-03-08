import pandas as pd
import xlsxwriter

from abc import ABC, abstractclassmethod
from typing import List, Dict, Tuple, Optional, Union

from .data_validation_config1_func import get_data_validation_dict,clean_df_data_validation
from .data_validation_typing import Header, SourceDict, SingleOptionsDict, DataValDict



class DataValidationConfiguration(ABC):

    @abstractclassmethod
    def __init__(self) -> None:
        self.data_validation_dict
        self.data_val_headers ## self.data_val_dict.keys()

    def set_data_validation(self, ws: xlsxwriter.worksheet.Worksheet, df: pd.DataFrame) -> None:
        
        """
        Set up data validation, dropdown lists 
        Parameters:
        ws: worksheet
        df: dataframe used to create the template header=None
        self.data_validation_dict: DataValDict Dictionary containing the opctions_dict for each field in scope for data validation
        self.data_validation_dict.keys(): List[Header] List of headers in scope for data validation
        """
        if self.data_validation_dict is None:
            return None

        column_indexes_to_apply_data_validation = [i for i, hd in enumerate(df.loc['HEADER']) if hd in self.data_val_headers]  
        initial_index = df.index.tolist().index('')  ## df index 0 = excel row 1
        last_row_index = df.shape[0] - 1  

        for col in column_indexes_to_apply_data_validation:
            hd = df.loc['HEADER', col]
            opts_dict = self.data_validation_dict[hd]
            ### ws.data_validation(first_row, first_col, last_row, last_col, options_dict={...})
            # ws.data_validation(initial_index, col, last_row_index, col, {'validate':'list', 'source':data_source_dict[hd], 'error_type':'stop'})
            ws.data_validation(initial_index, col, last_row_index, col, opts_dict)


class DataValidationConfig1(DataValidationConfiguration):
    """
    df_data_validation_complete: dataframe containing all the dropdown lists and the settings for the data validation
    df_data_validation: dataframe contaning only the dropdown lists 
    data_validation_dict: DataValDict

    """

    def __init__(self, df_dvconfig1: Union[pd.DataFrame,None], dropdown_list_sheet: str, df_settings: pd.DataFrame) -> None:
        if df_dvconfig1 is None:
            self.data_validation_dict = None
            self.data_val_headers = None
            self.df_data_validation_complete = None
            self.df_data_validation = None
        else:
            self.dropdown_list_sheet = dropdown_list_sheet
            self.df_data_validation_complete, self.df_data_validation = clean_df_data_validation(df_dvconfig1, df_settings)
            self.data_val_headers = self.df_data_validation.columns.tolist()
            self.data_validation_dict = get_data_validation_dict(df_settings, self.df_data_validation_complete, self.df_data_validation, self.dropdown_list_sheet)



