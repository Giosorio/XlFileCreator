



class HeaderIndexNotIdentified(Exception):
    def __init__(self, wrong_index: object) -> None:
        errormessage = f"""Index not identified: '{wrong_index}'
        Indexes accepted: 'CONFIG_MANAGER', 'header_format', 'lock_sheet_config', 'column_width', 'description_header', 'HEADER', 'example_row'"""
        print(errormessage)







