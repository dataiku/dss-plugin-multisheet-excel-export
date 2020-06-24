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

input_datasets_ids = get_input_names_for_role('input_dataset')
if len(input_datasets_ids) == 0:
    logger.warning("Received no input datasets ids. input_datasets_ids={}".format(input_datasets_ids))

input_datasets_names = [name.split('.')[-1] for name in input_datasets_ids]
if len(input_datasets_names) == 0:
    logger.warning("Received no input datasets names. input_datasets_ids={}, input_datasets_names={}".format(
        input_datasets_ids, input_datasets_names))

# Retrieve the list of output folders, should contain unique element
output_folder_id = get_output_names_for_role('folder')
logger.info("Retrieved the following folder ids: {}".format(output_folder_id))
output_folder_name = output_folder_id[0]
logger.info("Received the following output folder name: {}".format(output_folder_name))
output_folder = dataiku.Folder(output_folder_name)

input_config = get_recipe_config()
workbook_name = input_config.get('output_workbook_name', None)

if workbook_name is None:
    logger.warning("Received input received recipe config: {}".format(input_config))
    raise ValueError('Could not read the workbook name.')

if not str.isidentifier(workbook_name):
    raise ValueError("The input parameter workbook_name is not a valid identifier. "
                     "See the definition of an identifier at "
                     "https://docs.python.org/3/library/stdtypes.html?highlight=isidentifier#str.isidentifier")

excel_sheet_abs_path = os.path.join(output_folder.get_path(), '{}.xlsx'.format(workbook_name))
logger.info("Intend to write the output xls file to the following location: {}".format(excel_sheet_abs_path))

datasets_to_xlsx(input_datasets_names, excel_sheet_abs_path)
logger.info("Ended recipe processing.")