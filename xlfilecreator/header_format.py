import pandas as pd
import xlsxwriter

from typing import List, Union

from .formats import format_dict


def set_headers_format(wb: xlsxwriter.workbook.Workbook, ws: xlsxwriter.worksheet.Worksheet, 
df: pd.DataFrame, df_settings: pd.DataFrame, header_index_list: List, header_index: int) -> None:
    """
    Set format of the headers 

    worksheet.write(0, 0, 'Hello') -> Cell A1 = 'Hello'   header_index 0 = excel row 1   column 0 = excel column A

    Parameters:
    wb: workbook
    ws: worksheet
    df: data frame used to create the excel file
    df_settings: data frame containing the format settings, if there is no format specifications it will used format_0 as default (White backgorund and font in Bold)
    header_index_list: list of headers included in the index ['Description_header', 'HEADER', 'Example_header']
    """


    def set_format_hd(wb: xlsxwriter.workbook.Workbook, header_index: int, header_values: List, header_format: Union[List, str]):
        """
        wb: workbook object
        header_index: index where the values to format are located
        header_values: list of the headers in string format 
        header_format: List or string value of the format to apply, or list of string values of the formats to apply (string values must be part of the keys of format_dict)
        """
        
        # global format_dict
        
        if isinstance(header_format, str):    #### if not type(header_format) is list
            header_format = [header_format for i in df.columns]

        for col, header, hd_format in zip(df.columns, header_values, header_format):
            if hd_format == '':
                ws.write(header_index, col, header, wb.add_format(format_dict['format_0']))
            else:
                ws.write(header_index, col, header, wb.add_format(format_dict[hd_format]))


    # header_index = df.index.tolist().index('HEADER')
    header_values = df.loc['HEADER']
    header_format = df_settings.loc['header_format'].tolist()
    set_format_hd(wb, header_index, header_values, header_format)


    if 'example_row' in header_index_list:
        header_index = df.index.tolist().index('example_row')
        header_values = df.loc['example_row']
        set_format_hd(wb, header_index, header_values, 'format_10')


    if 'description_header' in header_index_list:
        header_index = df.index.tolist().index('description_header')
        header_values = df.loc['description_header']
        set_format_hd(wb, header_index, header_values, 'format_0')





