from setuptools import find_packages, setup


setup(
    name='xlfilecreator',
    packages=find_packages(include=['xlfilecreator']),
    version='0.1',
    description='Class XlFileTemp that splits a google sheet workbook on the basis of the values in one of the columns creating multiple password protected excel files',
    author='Giovanni Osorio',
    licence='MIT',
    install_requires=['pandas', 'openpyxl', 'xlsxwriter'],
    # test_suite='tests',
)