import pandas as pd

import glob
import os
import random
import string
from typing import Optional

from .utils_func import Project


class PackageMsofficeMissing(Exception):

    ### https://github.com/herumi/msoffice updates made on September 2023 do not work
    ### https://github.com/Giosorio/msoffice last update March 2023 work OK 
    errormessage = """
    Install msoffice to be able to encrypt excel files
    Check documentation here:
        https://github.com/Giosorio/msoffice

    Follow the steps to install msoffice from the terminal in your main directory
    Linux, Mac
        git clone https://github.com/Giosorio/cybozulib
        git clone https://github.com/Giosorio/msoffice
        cd msoffice
        make -j RELEASE=1
    
    Windows
        git clone https://github.com/herumi/cybozulib
        git clone https://github.com/herumi/msoffice
        git clone https://github.com/herumi/cybozulib_ext # for openssl
        cd msoffice
        mk.bat ; or open msoffice12.sln and build

    """


def _check_msoffice_installed(init: Optional[bool]=False) -> None:
    folders = glob.glob('*/')

    encrypt_folders = ['cybozulib/', 'msoffice/']
    if all(e_folder in folders for e_folder in encrypt_folders) is False:
        if init:
            print(f'\n{"*"*20}WARNING{"*"*20}')
            print(PackageMsofficeMissing.errormessage)
        else:
            raise PackageMsofficeMissing(PackageMsofficeMissing.errormessage)


def set_password(path_1: str, path_2: str, passwordMaster_name: str) -> None:

    def encrypt_file(password: str, path_in: str, path_out: str):
        """msoffice-crypt must be installed in the local folder"""

        os.system(f'msoffice/bin/msoffice-crypt.exe -e -p {password} {path_in} {path_out}')

    df_pw = pd.read_csv(passwordMaster_name)
    num_files = df_pw.shape[0]
    
    count = 1
    for file_n, pw in zip(df_pw['Filename'], df_pw['Password']):         
        path_in = '"{}/{}"'.format(path_1, file_n)
        path_out = '"{}/{}"'.format(path_2, file_n)
        encrypt_file(pw, path_in, path_out)
        # print(file_n, '   Password:', pw, f'---->{count}/{num_files}')
        count +=1
    

def create_password(project: Project, split_by_value: str, random_pw: Optional[bool]=False) -> str:
    """
    If random_pw is False is because there will be multiple batches and the password must remain the same
    password logic = project.name + str(123 * l) + split_by_value[:3][::-1]
    """

    _check_msoffice_installed()
    
    if random_pw is True:
        letters = string.ascii_uppercase + string.digits
        random_6 =  ''.join(random.choice(letters) for _ in range(6))
        if project.root == 'received':
            return project.name + random_6
        else:
            return random_6
        

    split_by_value = ''.join(char for char in split_by_value if char.isalnum())
    l = len(split_by_value)
    if project.root == 'received':
        pw = project.name + str(123 * l) + split_by_value[:3][::-1]  # split_by_value[:-4:-1]
    else:
        pw = str(123 * l) + split_by_value[:3][::-1]
    
    return pw.upper()



