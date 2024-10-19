import pandas as pd
from tqdm.auto import tqdm 

import datetime
from typing import Optional, List, Union, Dict

from .create_xlfile import process_template, protect_workbook
from .encrypt_xl import set_password, create_password
from .utils_func import set_project_name, create_output_folders, get_XlFile_details, password_dataframe, to_zip
from .xlfiletemp import XlFileTemp


def check_feasibility(template_list: List[XlFileTemp], split_by: str, split_by_range: List[str]) -> None:
    """
    All templates must have all split_value items provided in split_by_range list
    
    template_list: Python list containing the templates (XlFileTemp objects) to include in the Excel File.
    split_by: The name of the column to filter by.
    split_by_range: Python list contaning all the split_value items. If split_by_value=True All split_value items must be included in all templates provided.
    """
    for template in template_list:
        print(f"Checking: {template.tab_names['main_sheet']}")
        template.check_split_by_range(split_by, split_by_range)


def create_xl_file_multiple_temp(*, project_name: str, template_list: List[XlFileTemp], split_by_value: Union[bool,Dict[XlFileTemp,bool]], split_by: Optional[str]=None, 
    split_by_range: Optional[List[str]]=None, batch: Optional[int]=1, sheet_password: Optional[str]=None, workbook_password: Optional[str]=None,
    protect_files: Optional[bool]=False, random_password: Optional[bool]=False, in_zip: Optional[bool]=False) -> None:
    """
    Creates the Excel file with multiple tamples in it.

    project_name: name of the project, it will be part of the filename of the templates. If split_by is None it will be the name of the single file generated
    template_list: Python list containing the templates (XlFileTemp objects) to include in the Excel File.
    split_by_value: A boolean flag (True or False). If True, the method filters by the split_value provided. If False, it uses all values from the split_by column.
    split_by: The name of the column to filter by.
    split_by_range: Python list contaning all the split_value items. If split_by_value=True All split_value items must be included in all templates provided.
    batch: Number of the batch. Included in the filename of the templates.
    sheet_password: sheet password for the excel file to avoid the users to change the format of the main sheet, default=None 
    workbook_password: workbook password to avoid the users to add more sheets in the excel file, defaul=None
    protect_files: False/True encrypt the files
    random_password: False/True if protect_files is True it determines if the password of the files should be random or based on a logic
    in_zip: False/True Download folders in zip
    """

    if split_by is None and split_by_range is None:
        return None

    if isinstance(split_by_range, list):
        values_to_split = set(split_by_range)
    else:
        raise TypeError(f'{split_by_range} is not a list')

    ### Check feasibility
    if split_by_value is True:
        check_feasibility(template_list, split_by, split_by_range)

    ### Check that all templates provided in template_list are part of the split_by_value dictionary.keys() and the other way around 
    if isinstance(split_by_value, dict):
        if not all(template in template_list for template in split_by_value.keys()):
            raise ValueError('Not all templates provided in split_by_value are part of template_list')
        if not all(template in split_by_value.keys() for template in template_list):
            raise ValueError('Not all templates provided in template_list are part of split_by_value')
        if not all(isinstance(v, bool) for v in split_by_value.values()):
            raise ValueError(f'Invalid input split_by_value, only boolean values are accepted True/False. {split_by_value}')

        ### Check feasibility on templates where the split_by_value=True
        split_by_value_true_templates = [temp for temp, flag in split_by_value.items() if flag is True]
        check_feasibility(split_by_value_true_templates, split_by, split_by_range)
    
    ### Create output folders
    today = datetime.datetime.today().strftime('%Y%m%d')
    project = set_project_name(project_name)
    path_1, path_2 = create_output_folders(project.name, today, protect_files)
    
    ### 
    password_master = []
    pbar = tqdm(total=len(values_to_split))
    for i, split_value in enumerate(values_to_split, 1):
        
        ### Get Excelfile details (id, name, path)
        xl_file = get_XlFile_details(split_value, project, batch, i, today, path_1)
        
        ### Create Excel file
        with pd.ExcelWriter(xl_file.path, engine='xlsxwriter') as writer:

            for j, template in enumerate(template_list, 1):
                template_name = f'Sheet{j}'
                if isinstance(split_by_value, dict):
                    sbv = split_by_value[template]
                else:
                    sbv = split_by_value

                process_template(writer, template, sbv, template_name, split_by, split_value, sheet_password)
                
        ### Protect Workbook
        if workbook_password is not None and workbook_password != '':
            protect_workbook(xl_file.path, password=workbook_password)

        ### Create Password master df
        if protect_files is True:
            pw = create_password(project, split_value, random_password)    
            password_master.append((xl_file.id, xl_file.name, split_value, pw))

    ### Encrypt Excel files
    if protect_files is True:
        passwordMaster_name = password_dataframe(password_master, project, split_by, today)
        set_password(path_1, path_2, passwordMaster_name)

    if in_zip:
        to_zip(path_1, path_2)

    pbar.close()
