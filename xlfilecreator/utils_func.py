import pandas as pd

import collections
import datetime
import json
import os
from typing import List, Tuple, Optional, Union

from .data_validation_typing import DataValDict
from .terminal_colors import yellow
from .xlfilecreator_errors import HeaderIndexNotIdentified



def to_number(x):
	try:      
		x = float(x)
	except ValueError:
		pass
	
	return x


def validate_integer_input(x, source: str) -> int:
    if isinstance(x, bool):
        raise ValueError(f"Invalid integer input '{source}' --> {x}")

    try:
        x = int(x)
    except ValueError:
        raise ValueError(f"Invalid integer input '{source}' --> {x}")
    
    if x > 0:
        return x

    return 100
    

def get_column_to_split_by(df_settings: pd.DataFrame, split_by: str) -> int:
    """returns Dataframe integer column of the column to split by"""

    try:
        col_to_split = df_settings.loc['HEADER'].tolist().index(split_by)  
    except ValueError:
        raise ValueError(f"'{split_by}' not found in the HEADER")
    else:
        return col_to_split


def get_headers(df_settings: pd.DataFrame) -> Tuple[List[str], pd.DataFrame]:
    """
    Validate if all headers are included in the pre-stablished set of all_indexes
    If there is a value included in the index that is not part of accepted_idx it will raise an error

    Organise the headers according to headers_all
    'description_header' on the top
    'HEADER' second 
    'example_row' third
    """

    accepted_idx = ['CONFIG_MANAGER','header_format','lock_sheet_config','conditional_formatting','column_width','description_header','HEADER','example_row']
    for hd_i in df_settings.index:
        if hd_i not in accepted_idx:
            raise HeaderIndexNotIdentified(hd_i, accepted_idx)

    headers_all = ['description_header','HEADER','test_test_','example_row']
    header_index_list = [hd for hd in headers_all if hd in df_settings.index]
    df_hd = df_settings.loc[header_index_list]

    return header_index_list, df_hd


def rows_extra(df_data_only: pd.DataFrame, num_rows_extra: Optional[int]=100) -> pd.DataFrame:
    blank_rows = ['' for _ in range(num_rows_extra)]
    rows_extra = {col:blank_rows for col in df_data_only.columns}
    df_rows_extra = pd.DataFrame(rows_extra)

    return df_rows_extra


def get_df_data(df_hd: pd.DataFrame, df_data_only: pd.DataFrame, allow_input_extra_rows: Optional[bool]=False, num_rows_extra: Optional[int]=100) -> pd.DataFrame:
    """dataframe containing headers + data + extrarows"""

    if allow_input_extra_rows:
        df_rows_extra = rows_extra(df_data_only, num_rows_extra)
    else:
        df_rows_extra = None

    return pd.concat([df_hd, df_data_only, df_rows_extra])


def clean_df_main(df_main: pd.DataFrame) -> pd.DataFrame:
    """Remove blank columns"""

    columns_scope = [col for col in df_main.columns if df_main.loc['HEADER', col] != '']
    if len(columns_scope) == 0:
        errormessage = 'MAIN_SHEET is not readable. Possible Problems: CONFIG_MANAGER row does not have any values, or The HEADER row does not match any values in the IMPORT_FILE sheet'
        raise KeyError(errormessage)

    return df_main[columns_scope]


def get_excel_df(xl_file:str, sheet_name: str, header: Optional[str]=None) -> pd.DataFrame:
    """This function is only used to create the df_main or the df_dvconfig1"""
    df = pd.read_excel(xl_file, sheet_name=sheet_name, header=None, na_filter=False, index_col=0)
    df.index.name = 'Index'

    if header is None:
        df.columns = range(df.shape[1])
    else:
        try:
            df.columns = df.loc[header]
        except KeyError:
            raise KeyError(f"'HEADER' not found in the index {[idx for idx in df.index if idx != '']}")

    return df


def check_google_sh_reader(sheet_id: str, sheet_name: str, na_filter: bool, header: Union[int,None], index_col:Union[int,None]):
    """Check if the google sheet workbook is readeble"""
    try:
        df = pd.read_csv(f'https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}', na_filter=na_filter, header=header, index_col=index_col)
    except pd.errors.ParserError as pe:
        print(pe)
        raise pd.errors.ParserError('The Google sheet workbook is restricted. It must be accessible to Anyone with the link')
    else:
        return df


def get_google_sheet_df(sheet_id: str, sheet_name: str) -> pd.DataFrame:
    """Read google sheets main sheet"""
    df = check_google_sh_reader(sheet_id, sheet_name, na_filter=False, header=None, index_col=0)
    df.index.name = 'Index'
    
    df.columns = range(df.shape[1])
    return df


## data_validation_config1
def get_google_sheet_validation(sheet_id: str, dropdown_list_sheet: str) -> pd.DataFrame:
    """Read google sheets data_validation_config1"""
    if dropdown_list_sheet is None or dropdown_list_sheet == '':
        return None

    df = check_google_sh_reader(sheet_id, dropdown_list_sheet, na_filter=False, header=None, index_col=0)
    ## header: index 'HEADER' from the dropdows list sheet
    try:
        df.columns = df.loc['HEADER']
    except KeyError:
        print(yellow(f"\nWARNING: 'HEADER' not found in '{dropdown_list_sheet}', dropdown_list_sheet set as None\nEnsure '{dropdown_list_sheet}'is the correct name of the sheet and that it is in the correct format.\n"))
        return None
    else:
        df.index.name = 'Index'
        return df


## data_validation_config2
def get_google_sheet_validation2(sheet_id: str, data_validation_sheet_config2: str, dropdown_list_sheet: str) -> Tuple[pd.DataFrame,pd.DataFrame]:
    """
    Read google sheets data_validation_config2
    data_validation_sheet_config2: name of the sheet where the data_validation_config2 is located
    dropdown_list_sheet: name of the sheet where the dropdown lists are located
    """

    if data_validation_sheet_config2 is None or data_validation_sheet_config2 == '':
        return None, None
    if dropdown_list_sheet is None or dropdown_list_sheet == '':
        return None, None
        
    df_dvconfig2 = check_google_sh_reader(sheet_id, data_validation_sheet_config2, na_filter=False, header=0, index_col=None)
    df_picklists = check_google_sh_reader(sheet_id, dropdown_list_sheet, na_filter=False, header=0, index_col=None)
    return df_dvconfig2, df_picklists


def get_excel_dvalidation2(xl_file:str, data_validation_sheet_config2: str, dropdown_list_sheet: str) -> Tuple[pd.DataFrame,pd.DataFrame]:
    """
    Read excel file data_validation_config2
    data_validation_sheet_config2: name of the sheet where the data_validation_config2 is located
    dropdown_list_sheet: name of the sheet where the dropdown lists are located
    """

    if data_validation_sheet_config2 is None or data_validation_sheet_config2 == '':
        return None, None
    if dropdown_list_sheet is None or dropdown_list_sheet == '':
        return None, None
        
    df_dvconfig2 = pd.read_excel(xl_file, sheet_name=data_validation_sheet_config2, na_filter=False)
    df_picklists = pd.read_excel(xl_file, sheet_name=dropdown_list_sheet, na_filter=False)
    return df_dvconfig2, df_picklists


def export_json(dict_: DataValDict, filename: str) -> None:
    json_str_format = json.dumps(dict_, indent=2)   ## string
    filename = f'{filename}.json'
    with open(filename, 'w') as outfile:
        outfile.write(json_str_format)


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