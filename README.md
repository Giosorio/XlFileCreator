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

### Main configuration
```python

### Tab names
main_sheet = 'MAIN_SHEET'
data_validation_sheet_config1 = 'data_validation_sheet_config1'
data_validation_sheet_config2 = 'data_validation_sheet_config2'
dropdown_lists_sheet_config2 = 'dropdown_lists_sheet_config2'
conditional_formatting_sheet = 'conditional_formatting_sheet'

project_name = 'TEST'

### Name of the column to filter and create new templates. If split_by_range is provided the data is not filtered by the values in the dataset.
### It will replicate the data from the main dataset for each unique value in split_by_range.
split_by = 'Supplier'

### split_by_range: A Python list containing values to split the data by. It is used to create a separate template for each unique value in the list.
split_by_range = None

### batch: Number of the batch. Included in the filename of the templates 
batch = 1

### Optional
sheet_password = '123'
workbook_password = '456'
allow_input_extra_rows = True

### Number of extra empty rows. Ignored if allow_input_extra_rows=False. If allow_input_extra_rows=True the default value is 100
num_rows_extra = 10000

### Optional. 
protect_files = True
random_password = False
in_zip = True

### identify_data_types: Convert the numbers read as text into float values
### Passing identify_data_types=False can improve the performance of reading a large file and numbers will remain in text format
identify_data_types = False

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
### Google sheet id
sheet_id = '1UVR_tnuLJwqag45dweRBGjQbxQGxW3cXCfyazi4AQEw'

project_1 = XlFileTemp.read_google_sheets_file(sheet_id, main_sheet, data_validation_sheet_config1,
            data_validation_sheet_config2, dropdown_lists_sheet_config2, conditional_formatting_sheet, identify_data_types)

project_1.to_excel(project_name, split_by, split_by_range, batch, sheet_password, workbook_password, allow_input_extra_rows, num_rows_extra, protect_files, random_password, in_zip)

```

### Config file read from Excel
* The Config file is commented to guide you through setting up the template's configuration.


```python
### Path of the Excel file
xl_file = 'XlFileTemp_config_file.xlsx'

project_1 = XlFileTemp.read_excel(xl_file, main_sheet, data_validation_sheet_config1,
        data_validation_sheet_config2, dropdown_lists_sheet_config2,
        conditional_formatting_sheet, identify_data_types)

project_1.to_excel(project_name, split_by, split_by_range, batch, sheet_password, workbook_password, allow_input_extra_rows, num_rows_extra, protect_files, random_password, in_zip)

```






























                   


