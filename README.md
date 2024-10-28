# XlFileCreator
This Python package automates the creation of Excel templates based on a configuration provided in a Google Sheets or Excel input file. 

It empowers you to:

* **Generate numerous Excel templates effortlessly:** Imagine needing 200 individual data collection Excel files - this tool handles that, creating them with ease.
* **Personalise templates with pre-filled data:** Add unique information for each recipient, such as worker details for specific suppliers.
* **Apply advanced formatting:**  Automate formatting tasks like conditional formatting, dropdown lists, formulas, data types, and header styling.
* **Secure your templates:** Protect sheets and workbooks, and encrypt your files for enhanced security. 

This tool saves significant time by creating multiple Excel templates with uniform configurations, tailored content, and built-in validations—ideal for large-scale data collection or reporting tasks.

### Customisable features
* **Conditional Formatting**
* **Dropdown Lists**
* **Formulas**
* **Header Styling**
* **Data Types (Date, Currency, Number, General)**
* **Sheet and Workbook Protection**
* **File Encryption**


Installation
------------

##### Requirements
* Python     (Tested using 3.9 and 3.10)
* Pandas     (Tested using 1.4.4)
* Xlswriter  (Tested using 3.0.3)
* Openpyxl   (Tested using 3.0.10)
* [herumi/msoffice](https://github.com/herumi/msoffice) (file encryption)


##### COLAB example (linux) herumi/msoffice from GitHub
    !git clone https://github.com/herumi/cybozulib
    !git clone https://github.com/herumi/msoffice
    !cd msoffice; make -j RELEASE=1
    
##### XlFileCreator from GitHub
    pip install git+https://github.com/Giosorio/XlFileCreator@[VERSION TAG]


How to use
------------

```python
from xlfilecreator.xlfiletemp import XlFileTemp

### Get an Excel Config file:
XlFileTemp.export_config_file()


```

### Config file read from Google Sheets
#### Before running the script:

1.   The Config file is commented to guide you through setting up the template's configuration.
2.   The Google sheet file must be visible to anyone with the link.
3.   The data must be imported from a CSV file in plain text (when importing the file in Google Sheets **UNTICK** the box 'Convert text to numbers, dates and formulas')
4.   If filters are applied in any of the sheets part of the process, the script will read the filtered data as shown in the Google Sheets.
Suggestion: Remove the filters to run all data.
5. Delete all extra columns that are not used from the MAIN_SHEET.
6. Ensure that the formulas in the criteria in the data_validation_sheet_config2 and conditional_formatting_sheet are visible using a ' at the beginning of the formula. Any other character will throw an error.

```python

template_1 = XlFileTemp.read_google_sheets_file(
            sheet_id='1pBqN6b4HyfKo_SJuSndORasDJXTpTWB3Dwj8fcrYxJ4',     # Google sheet id
            main_sheet='MAIN_SHEET', 
            data_validation_sheet_config1='data_validation_config1',     # Optional[str]=None
            data_validation_sheet_config2='data_validation_config2',     # Optional[str]=None
            dropdown_lists_sheet_config2='dropdown_lists_config2',       # Optional[str]=None
            conditional_formatting_sheet='conditional_formatting',       # Optional[str]=None
            identify_data_types= True,        ## Optional[bool]=True Convert the numbers read as text into float values
            )

```

### Config file read from Excel
* The Config file is commented to guide you through setting up the template's configuration.


```python
### Path of the Excel file

template_1 = XlFileTemp.read_excel(
        xl_file='XlFileTemp_config_file_TEST.xlsx',                  # Excel File path
        main_sheet='MAIN_SHEET', 
        data_validation_sheet_config1='data_validation_config1',     # Optional[str]=None
        data_validation_sheet_config2='data_validation_config2',     # Optional[str]=None
        dropdown_lists_sheet_config2='dropdown_lists_config2',       # Optional[str]=None
        conditional_formatting_sheet='conditional_formatting',       # Optional[str]=None
        identify_data_types= True,            ## Optional[bool]=True Convert the numbers read as text into float values
        )


```

## Generating Excel Files with a Single Template

#### Parameters:
* **project_name:** Optional[str]=None name of the project, it will be part of the filename of the templates. If split_by is None it will be the name of the single file generated
* **split_by:** Optional[str]=None Name of the column used to filter and create new templates. If split_by_range is provided, the data is not filtered by this column's values. Instead, the data is duplicated for each unique value in split_by_range.
* **split_by_range:** Optional[List[str]]=None A Python list containing values to split the data by. It is used to create a separate template for each unique value in the list.
* **batch:** Optional[int]=1 Number of the batch. Included in the filename of the templates 
* **sheet_password:** Optional[str]=None sheet password for the excel file to avoid the users to change the format of the main sheet, default=None 
* **workbook_password:** Optional[str]=None workbook password to avoid the users to add more sheets in the excel file, defaul=None
* **allow_input_extra_rows:** Optional[bool]=None False/True Determines if the templates allow the user to fill out more rows in the template, If None it will use self.extrarows
* **num_rows_extra:** Optional[int]=None Number of additional rows to include in the template. If set to None and allow_input_extra_rows=True, the default is 100. Ignored if allow_input_extra_rows=False
* **protect_files:** Optional[bool]=False False/True encrypt the files
* **random_password:** Optional[bool]=False if protect_files is True it determines if the password of the files should be random or based on a logic
* **in_zip:** Optional[bool]=False Download folders in zip 




### Template preview: 
It will create a single Excel file including the formatting, data validation and conditional formatting

```python

template_1.to_excel(
        project_name='EXCELTEST',   # Optional[str]=None
    )


```



### Create templates filtering by a specific column in the dataset
**split_by =** 'Supplier' Header of the column to split by

```python

template_1.to_excel(
        project_name='PROJECT',         # Optional[str]=None
        split_by='Supplier',            # Optional[str]=None
        batch=1,                        # Optional[int]=1
        sheet_password='123',           # Optional[str]=None
        workbook_password='456',        # Optional[str]=None
        protect_files=True,             # Optional[bool]=False
        random_password=False,          # Optional[bool]=False
        in_zip=True,                    # Optional[bool]=False
        allow_input_extra_rows=True,    # Optional[bool]=None
        num_rows_extra=10000,           # Optional[int]=None   Ignored if allow_input_extra_rows=False
        )

```



### Create templates duplicating the original dataset and setting each unique value from the split_by_range list provided.
In the following example, it will create 3 Excel templates and it will set each value provided in the split_by_value list in the entire supplier column.

```python

template_1.to_excel(
        project_name='PROJECT',         # Optional[str]=None
        split_by='Supplier',            # Optional[str]=None
        split_by_range=['AAA','BBB','CCC'],   # Optional[List[str]]=None
        batch=1,                        # Optional[int]=1
        sheet_password='123',           # Optional[str]=None
        workbook_password='456',        # Optional[str]=None
        protect_files=True,             # Optional[bool]=False
        random_password=False,          # Optional[bool]=False
        in_zip=True,                    # Optional[bool]=False
        allow_input_extra_rows=True,    # Optional[bool]=None
        num_rows_extra=10000,           # Optional[int]=None   Ignored if allow_input_extra_rows=False
        )

```


## Generating Excel Files with Multiple Templates

```python
from xlfilecreator.xlfiletemp import create_xl_file_multiple_temp

```

#### Key Parameters:

* **template_list:** List[XlFileTemp] Python list containing the templates (XlFileTemp objects) to include in the Excel File.
* **split_by_value:** Union[bool,Dict[XlFileTemp,bool]] A boolean flag (True or False) Or Dictionary {Temp: bool}. If True, the method filters by the split_value provided. If False, it uses all values from the split_by column.
* **split_by:** Optional[str]=None The name of the column to filter by.
* **split_by_range:** Optional[List[str]]=None Python list contaning all the split_value items. **If split_by_value=True All split_value items must be included in all templates provided.**

### Option 1
Creates three Excel file templates, one for each value in the split_by_range list. Each file will contain two tabs, one for each template. All three values in split_by_range must appear under the same column header, split_by='Supplier', in both templates from template_list.

```python

create_xl_file_multiple_temp(
    project_name='ABCD', 
    template_list=[template_1,template_2], 
    split_by_value=True,
    split_by='Supplier', 
    split_by_range=['AAA','BBB','CCC'],
    sheet_password='123', 
    workbook_password='456',
    protect_files=True,
    in_zip=False,
    )

```

### Option 2

Creates three Excel file templates, one for each value in split_by_range. Each file contains two tabs, one for each template in template_list. In templates where split_by_value is True, all values in split_by_range must appear under the column header specified by split_by='Supplier'. When split_by_value is False for a template, the full dataset from that template is replicated in the Excel file, with the split_by column ('Supplier') set to the corresponding value from split_by_range throughout.

**In conclusion:**

template_1 will be filtered by each supplier in the dataset, creating 3 Excel files for each value in split_by_range.
template_2 will contain the full dataset in each of the 3 Excel files, with the split_by column ('Supplier') set to each supplier name across the entire column in each file.

```python

create_xl_file_multiple_temp(
    project_name='ABCD', 
    template_list=[template_1,template_2], 
    split_by_value={template_1: True, template_2: False},
    split_by='Supplier', 
    split_by_range=['AAA','BBB','CCC'],
    sheet_password='123', 
    workbook_password='456',
    protect_files=True,
    in_zip=False,
    )

```





















                   


