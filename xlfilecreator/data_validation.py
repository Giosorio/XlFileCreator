import pandas as pd
import xlsxwriter

from typing import List, Dict, Tuple, Optional, Union

from .utils_func import export_json


Header = str # 'Worker Type'
Source = str # '=Droptdownlist!$F$2:$F$3'
SourceDict = Dict[Header, Source]
            # {
                # 'Worker Gender': '=Droptdownlist!$B$2:$B$4',
                # 'Worker Pay Type Name': '=Droptdownlist!$C$2:$C$5',
                # 'Rate type': '=Droptdownlist!$F$2:$F$3'
            # }
SingleOptionsDict = Dict[Header, Dict[str,str]] 
                # {'Worker Type':{
                                # 'validate': 'list',
                                #   'source': '=Droptdownlist!$F$2:$F$3',
                                #   'error_type': 'warning',
                                #   'input_title': 'Worker Paytype',
                                #   'input_message': 'Select a value from the picklist',
                                #   'error_title': 'Input value not valid!',
                                #   'error_message': 'It should be a value from the picklist'
                                #   }
                        # }
DataValDict = Dict[Header,SingleOptionsDict]
            # {
                # 'Worker Gender': 
                        # {'validate': 'list',
                        #   'source': '=Dropdown_Lists!$A$2:$A$4',
                        #   'error_type': 'stop',
                        #   'error_title': 'Worker Gender ERROR TITLE'},
                #  'Worker Pay Type Name': 
                        # {'validate': 'list',
                        #   'source': '=Dropdown_Lists!$B$2:$B$5',
                        #   'error_type': 'information'},
                #  'ACTIVE WORKER?': 
                        # {'validate': 'list',
                        #   'source': '=Dropdown_Lists!$C$2:$C$3',
                        #   'error_type': 'stop',
                        #   'input_title': 'Active worker?',
                        #   'input_message': 'Is this worker an active employee?',
                        #   'error_message': 'it must be an option from the picklist'
            #  }



def set_data_validation(ws: xlsxwriter.worksheet.Worksheet, df: pd.DataFrame, 
data_validation_opts_dict: Dict[str,Dict], data_val_headers: List[str]) -> None:
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


def get_data_validation_sources_dict(df_settings: pd.DataFrame, df_data_validation: pd.DataFrame, 
dropdown_list_sheet: str) -> Tuple[SourceDict, List[Header]]:
    """
    Return a dictionary and a list
    - The dictionary where the Keys are the headers to apply data validation and the Values are 
        the string formats of the excel range where the data validation is located
        
        It assumes that all the ranges start from row 2 in excel  
        SourceDict: 
            data_source_dict = {
                'Worker Gender': '=Droptdownlist!$B$2:$B$4',
                'Worker Pay Type Name': '=Droptdownlist!$C$2:$C$5',
                'Rate type': '=Droptdownlist!$F$2:$F$3'
            }

    - The list of the headers in to apply data validation.
        i.e ['Worker Gender','Worker Pay Type Name','ACTIVE WORKER?','Rate type']
    
    """

    data_source_dict = {}

    data_val_headers = df_data_validation.columns.tolist()
    ### Removing Unnamed columns
    data_val_headers = [hd for hd in data_val_headers if hd in df_settings.loc['HEADER'].tolist()]

    for col_num, hd in enumerate(data_val_headers):
        col_letter = xlsxwriter.utility.xl_col_to_name(col_num)
        last_row_index = len(df_data_validation[hd][df_data_validation[hd]!='']) + 1
        source_format = f'={dropdown_list_sheet}!${col_letter}$2:${col_letter}${last_row_index}'
        data_source_dict[hd] = source_format

    return data_source_dict, data_val_headers


def get_options_dict_data_validation(hd: str, source: str, opts_dv_included: List, 
df_data_validation_complete: pd.DataFrame) -> SingleOptionsDict:
    """
    Creates the options_dictionary for data validation FOR EACH HEADER
    The header as a key and the options dictionary as the value 
    
    worksheet.data_validation('B27', options_dict={'validate': 'list',
                                  'source': '=Droptdownlist!$F$2:$F$3',
                                  'error_type': 'warning',
                                  'input_title': 'Worker Paytype',
                                  'input_message': 'Select a value from the picklist',
                                  'error_title': 'Input value not valid!',
                                  'error_message': 'It should be a value from the picklist'})
    """

    options_dict = {'validate':'list', 'source':source, 'error_type':'stop'}
    for opt in opts_dv_included:
        opt_value = df_data_validation_complete.loc[opt, hd]
        if opt_value == '':
            continue 
        else:
            options_dict[opt]= opt_value

    return options_dict


def get_data_validation_dict(df_settings: pd.DataFrame, df_data_validation_complete: pd.DataFrame, 
df_data_validation: pd.DataFrame, dropdown_list_sheet: Optional[str]='Dropdown_Lists') -> Union[None, Tuple[DataValDict, List[Header]]]:
    """
    Generate a dictionary where the keys are the headers to apply data validation 
    and the values are the dictionaries containing the options for the data validation 
    {
        'Worker Paytype': {'validate':'list', 'source':'=Droptdownlist!$F$2:$F$3', 'error_type':'stop', 'input_title':'', 'input_message':'', 'error_title':'', 'error_message':'',},
        'header_2': {'validate':'list', 'source':'=Droptdownlist!$C$2:$C$12', 'error_type':'warning', 'input_title':'', 'input_message':'', 'error_title':'', 'error_message':'',},
    }
    
    Parameters:
    df_settings: dataframe used to create the template header=None
    df_data_validation_complete: dataframe containing all settings for data validation
    df_data_validation: dataframe containing only the dropdownlists 
    """

    if df_data_validation is None:
        return None

    data_source_dict, data_val_headers = get_data_validation_sources_dict(df_settings, df_data_validation, dropdown_list_sheet)
    # print(data_source_dict)
    options_dv_all = ['error_type', 'input_title', 'input_message', 'error_title', 'error_message']
    opts_dv_included = [opt for opt in options_dv_all if opt in df_data_validation_complete.index]
    
    data_validation_opts_dict = {}
    for hd in data_val_headers:
        source = data_source_dict[hd]
        opts_dict = get_options_dict_data_validation(hd, source, opts_dv_included, df_data_validation_complete)
        data_validation_opts_dict[hd] = opts_dict
        # ws.data_validation(initial_index, col, last_row_index, col, opts_dict)

    export_json(data_validation_opts_dict, 'data_validation_settings')
    
    return data_validation_opts_dict, data_val_headers


def clean_df_data_validation(df_data_validation_complete: pd.DataFrame, df_settings: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:

    if df_data_validation_complete is None:
        return None, None

    HEADER = [hd for hd in df_settings.loc['HEADER'] if hd != '']
    DV_HEADER = [hd for hd in df_data_validation_complete.loc['HEADER'] if hd != '']
    hd_included = [hd for hd in DV_HEADER if hd in HEADER]
    df_data_validation_complete = df_data_validation_complete[hd_included]
    df_data_validation = df_data_validation_complete[df_data_validation_complete.index=='']
    df_data_validation.columns = df_data_validation_complete.loc['HEADER']

    return df_data_validation_complete, df_data_validation
