import pandas as pd
from xlsxwriter.utility import xl_col_to_name

from typing import Optional, Dict, List

from .data_validation_typing import DataValDict, Header


# def config2(df_config2: pd.DataFrame, df_settings: pd.Dataframe)->pd.Dataframe:

#     header_values = df_settings.loc['HEADER'].tolist()
#     apply_to_xlcolumn = []

#     config_rows = df_config2.shape[0]
#     for row in range(config_rows):
#         hd_to_apply = df_config2.loc[row, 'apply_to']
#         idx_number_col = header_values.index(hd_to_apply)
#         apply_to_xlcolumn.append(xl_col_to_name(idx_number_col))

#     df_config2['apply_to_xlcolumn'] = apply_to_xlcolumn


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


def get_data_validation_dict_config2(df_dvconfig2: pd.DataFrame)->DataValDict:

    data_validation_opts_dict = {}
    config_rows = df_dvconfig2.shape[0]
    for row in range(config_rows):
        hd_to_apply = df_dvconfig2.loc[row, 'apply_to']
        opts_dict = create_opts_dict(df_dvconfig2.loc[row])
        data_validation_opts_dict[hd_to_apply] = opts_dict

    return data_validation_opts_dict

