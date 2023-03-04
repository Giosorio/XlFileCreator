from typing import List



class HeaderIndexNotIdentified(Exception):
    def __init__(self, wrong_index: str, accepted_idx: List) -> None:
        errormessage = f"""Index not identified: '{wrong_index}'\nIndexes accepted: {accepted_idx}"""
        print(errormessage)







