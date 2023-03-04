import pandas as pd
import xlsxwriter

from typing import List

from .formats import format_dict


def highlight_mandatory(wb: xlsxwriter.workbook.Workbook, ws: xlsxwriter.worksheet.Worksheet, 
df: pd.DataFrame, df_settings: pd.DataFrame, data_index: int) -> None:
    """
    Highlight mandatory fields in yellow
    
    wb: workbook
    ws: worksheet
    df: data frame used to create the excel file
    df_settings: data frame containing the format settings, if there is no format specifications it will used format_0 as default (White backgorund and font in Bold)
    data_index: interger index where the data starts in the df
    """
    
    if 'conditional_formatting' not in df_settings.index:
        return None

    length = df.shape[0]
    cond_formatting = df_settings.loc['conditional_formatting']
    for col_num, cond_f in zip(df.columns, cond_formatting):
        if cond_f == 'Mandatory':
            col_letter = xlsxwriter.utility.xl_col_to_name(col_num)
            ws.conditional_format(f'{col_letter}{data_index+1}:{col_letter}{length}',{'type': 'formula', 'criteria': f'=${col_letter}{data_index+1}=""', 'format': eval(format_dict['format_12'])})



# def set_conditional_formatting():
#     highlight_mandatory()

