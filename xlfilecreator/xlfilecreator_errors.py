from typing import List



class HeaderIndexNotIdentified(Exception):
    def __init__(self, wrong_index: str, accepted_idx: List[str]) -> None:
        errormessage = f"""Index not identified: '{wrong_index}'\nIndexes accepted: {accepted_idx}"""
        print(errormessage)


class MainSheetNoData(Exception):
    def __init__(self, *args: object) -> None:
        errormessage = 'MAIN_SHEET is not readable\nPossible Problems:\nCONFIG_MANAGER row does not have any values\nThe HEADER row does not match any values in the IMPORT_FILE sheet\n'
        print(errormessage)



