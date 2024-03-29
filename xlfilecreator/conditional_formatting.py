import pandas as pd
import xlsxwriter

from typing import Dict

from .formats import format_dict
from .terminal_colors import yellow


def highlight_mandatory(wb: xlsxwriter.workbook.Workbook, ws: xlsxwriter.worksheet.Worksheet, 
df: pd.DataFrame, df_settings: pd.DataFrame, data_index: int, allow_input_extra_rows: bool, num_rows_extra: int) -> None:
    """
    Highlight mandatory fields in yellow
    If allow_input_extra_rows is True, the extra rows will not have conditional formatting.
    
    wb: workbook
    ws: worksheet
    df: data frame used to create the excel file
    df_settings: data frame containing the format settings, if there is no format specifications it will used format_0 as default (White backgorund and font in Bold)
    data_index: interger index where the data starts in the df
    """
    
    if 'conditional_formatting' not in df_settings.index:
        return None

    if allow_input_extra_rows:
        length = df.shape[0] - num_rows_extra
    else:
        length = df.shape[0]

    cond_formatting = df_settings.loc['conditional_formatting']
    for col_num, cond_f in zip(df.columns, cond_formatting):
        if cond_f == 'Mandatory':
            col_letter = xlsxwriter.utility.xl_col_to_name(col_num)
            ws.conditional_format(f'{col_letter}{data_index+1}:{col_letter}{length}',{'type': 'formula', 'criteria': f'=${col_letter}{data_index+1}=""', 'format': wb.add_format(format_dict['format_12'])})



class CondFormatting:

    def __init__(self, df_condf: pd.DataFrame, df: pd.DataFrame) -> None:
        self.df_condf = CondFormatting.df_condf_validation(df_condf, df)

    @staticmethod
    def df_condf_validation(df_condf: pd.DataFrame, df: pd.DataFrame)->pd.DataFrame:
        
        if df_condf is None:
            return None

        try:
            df_condf['valid_apply_to'] = [hd in df.loc['HEADER'].tolist() for hd in df_condf['apply_to']]
            ### Valid type should be validated against a list of accepted values for future versions
            df_condf['valid_type'] = [t != '' for t in df_condf['type']]
            df_condf['valid_criteria'] = [c != '' for c in df_condf['criteria']]
            df_condf['valid_format'] = [f in format_dict.keys() for f in df_condf['format']]
            
            ### Overall Validation
            df_condf['overall_validation'] = [all((v_apply_to, v_type, v_criteria, v_format)) for v_apply_to, v_type, v_criteria, v_format in zip(df_condf['valid_apply_to'], df_condf['valid_type'], df_condf['valid_criteria'], df_condf['valid_format'])]
            df_condf = df_condf[df_condf['overall_validation']==True]
        except KeyError as ke:
            print(yellow(f"\nWARNING: Column {ke} not found in the conditional formatting sheet.\nconditional_formatting set as None\n"))
            return None

        return df_condf

    @staticmethod
    def create_opts_dict(wb, opts_settings:pd.Series) -> Dict[str,str]:
        format_opts_dict = format_dict[opts_settings['format']]

        opts_dict={'type': opts_settings['type'],
            'criteria': opts_settings['criteria'],
            'format': wb.add_format(format_opts_dict)}

        return opts_dict

    def set_conditional_formatting(self, wb: xlsxwriter.workbook.Workbook, ws: xlsxwriter.worksheet.Worksheet, df: pd.DataFrame) -> None:
        
        """
        Set up conditional formatting
        Parameters:
        ws: worksheet
        df: dataframe used to create the template header=None
        self.df_condf: dataframe containing the configuration for the data validation
        """

        if self.df_condf is None:
            return None

        initial_index = df.index.tolist().index('')  ## df index 0 = excel row 1
        last_row_index = df.shape[0] - 1
        header_list = df.loc['HEADER'].tolist()

        for row in self.df_condf.index:
            hd_apply_to = self.df_condf.loc[row, 'apply_to']
            col_idx = header_list.index(hd_apply_to)
            opts_dict = CondFormatting.create_opts_dict(wb, self.df_condf.loc[row])
            ws.conditional_format(initial_index, col_idx, last_row_index, col_idx, opts_dict)
