"""
Custom recipe for Excel Multi Sheet Exporter
"""

import pandas as pd
import dataiku
import logging
from dataiku.customrecipe import *

from utils import datasets_to_xlsx

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format='Multi-Sheet Excel Exporter | %(levelname)s - %(message)s')

input_datasets_ids = get_input_names_for_role('dataset')
input_datasets_names = [name.split('.')[-1] for name in input_datasets_ids]

# Retrieve the list of output folders, should contain unique element
output_folder_id = get_output_names_for_role('folder')
output_folder_name = output_folder_id[0]
output_folder = dataiku.Folder(output_folder_name)

input_config = get_recipe_config()
workbook_name = input_config.get('output_workbook_name')
if not str.isidentifier(workbook_name):
    raise ValueError("The input parameter workbook_name is not a valid identifier. "
                     "See the definition of an identifier at "
                     "https://docs.python.org/3/library/stdtypes.html?highlight=isidentifier#str.isidentifier")

excel_sheet_abs_path = os.path.join(output_folder.get_path(), '{}.xlsx'.format(workbook_name))

datasets_to_xlsx(input_datasets_names, excel_sheet_abs_path)
