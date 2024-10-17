from setuptools import find_packages, setup


setup(
    name='xlfilecreator',
    packages=find_packages(include=['xlfilecreator','tqdm']),
    version='0.41',
    description='Class XlFileTemp that splits a google sheet workbook on the basis of the values in one of the columns creating multiple password protected excel files, It includes dropdown lists, and conditional formatting',
    author='Giovanni Osorio',
    licence='MIT',
    install_requires=['pandas', 'openpyxl', 'xlsxwriter', 'tqdm'],
)