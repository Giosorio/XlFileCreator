import pandas as pd

import collections
import datetime
import json
import os
import string
from typing import List, Tuple, Dict, Optional

from .terminal_colors import yellow
from .xlfilecreator_errors import HeaderIndexNotIdentified



def get_column_to_split_by(df_settings, split_by):
    """returns Dataframe integer column of the column to split by"""

    try:
        col_to_split = df_settings.loc['HEADER'].tolist().index(split_by)  
    except ValueError:
        raise ValueError(f"'{split_by}' not found in the HEADER")
    else:
        return col_to_split


def get_headers(df_settings: pd.DataFrame) -> Tuple[List, pd.DataFrame]:
    """
    Validate if all headers are included in the pre-stablished set of all_indexes
    If there is a value included in the index that is not part of all_indexes it will raise an error

    Organise the headers according to headers_all
    'description_header' on the top
    'HEADER' second 
    'example_row' third
    """

    all_indexes = ['CONFIG_MANAGER','header_format','lock_sheet_config','column_width','description_header','HEADER','example_row']
    for hd_i in df_settings.index:
        if hd_i not in all_indexes:
            raise HeaderIndexNotIdentified(hd_i)

    headers_all = ['description_header','HEADER','test_test_','example_row']
    header_index_list = [hd for hd in headers_all if hd in df_settings.index]
    df_hd = df_settings.loc[header_index_list]

    return header_index_list, df_hd


def rows_extra(df_data_only: pd.DataFrame) -> pd.DataFrame:
    blank_rows = ['' for _ in range(100)]
    rows_extra = {col:blank_rows for col in df_data_only.columns}
    df_rows_extra = pd.DataFrame(rows_extra)

    return df_rows_extra


def get_df_data(df_hd: pd.DataFrame, df_data_only: pd.DataFrame, allow_input_extra_rows: Optional[bool]=False) -> pd.DataFrame:
    """dataframe containing headers + data + extrarows"""

    if allow_input_extra_rows:
        df_rows_extra = rows_extra(df_data_only)
    else:
        df_rows_extra = None
        ### in case 'allow_input_extra_rows' has changed more than once
        df_data_only = df_data_only[df_data_only[0]!='']

    return pd.concat([df_hd, df_data_only, df_rows_extra])


def clean_df_main(df_main: pd.DataFrame) -> pd.DataFrame:
    """Remove blank columns"""

    columns_scope = [col for col in df_main.columns if df_main.loc['HEADER', col] != '']
    
    return df_main[columns_scope]


def get_google_sheet_df(sheet_id: str, sheet_name: str, header: Optional[str]=None) -> pd.DataFrame:
    """header: index 'HEADER' from the dropdows list sheet"""
    ### Read google sheets file
    try:
        df = pd.read_csv(f'https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}', na_filter=False, header=None, index_col=0)
    except pd.errors.ParserError as pe:
        print(pe)
        raise pd.errors.ParserError('The Google sheet workbook is restricted. It must be accessible to Anyone with the link')

    df.index.name = 'Index'

    if header is None:
        df.columns = range(df.shape[1])
    else:
        df.columns = df.loc[header]

    return df


def get_google_sheet_validation(sheet_id, dropdown_list_sheet):
    try:
        df = get_google_sheet_df(sheet_id, dropdown_list_sheet, header='HEADER')
    except KeyError:
        print(yellow(f"\nWARNING: 'HEADER' not found in '{dropdown_list_sheet}', dropdown_list_sheet set as None\nEnsure '{dropdown_list_sheet}'is the correct name of the sheet and that it is in the correct format.\n"))
        return None
    else:
        return df


def export_json(dict_: Dict, filename: str) -> None:
    json_str_format = json.dumps(dict_, indent=2)   ## string
    filename = f'{filename}.json'
    with open(filename, 'w') as outfile:
        outfile.write(json_str_format)
    
    ### google colab:
    # files.download(filename)

Project = Tuple[str, str]
def set_project_name(project_name: Optional[str]=None) -> Project:
    ProName = collections.namedtuple('ProName', ['name', 'root'])

    if project_name is None or project_name == '':
        now = datetime.datetime.now().strftime('%Y%m%d_%H-%M-%S')
        project = ProName(name=f'Project_{now}', root='default')
    else:
        project_name = ''.join(char for char in project_name if char.isalnum())
        project = ProName(name=project_name, root='received')
    
    return project


def create_output_folders(project_name: str, today: str, protect_files: Optional[bool]=False) -> Tuple[str, str]:

    path_1 = f'{project_name}_XL_files_{today}'
    path_2 = f'{project_name}_XL_files_password_{today}'
    os.mkdir(path_1)

    if protect_files:
        os.mkdir(path_2)
    
    return path_1, path_2