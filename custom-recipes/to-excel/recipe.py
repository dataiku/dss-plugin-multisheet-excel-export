import logging
from dataiku import pandasutils as pdu
import pandas as pd
import re
import dataiku
import dataikuapi
from dataiku import api_client as client
from dataiku.core.sql import SQLExecutor2
import json, time, logging
from dataiku.customrecipe import *

logger = logging.getLogger(__name__)

# Prepare and check parameters

output_first_names = get_output_names_for_role('folder')
output_first_name = output_first_names[0]
output_folder = dataiku.Folder(output_first_names[0])
path = output_folder.get_path()

input_names = get_input_names_for_role('dataset')
sheet_names = get_recipe_config()['sheet_names']

def get_input_dataset(role, i):
    names = get_input_names_for_role(role)
    return dataiku.Dataset(names[i]) if len(names) > 0 else None

workbook_name = get_recipe_config()['output_workbook_name']

writer = pd.ExcelWriter(path + '/' + workbook_name + '.xlsx', engine='openpyxl')

# Write datasets to workbook

def datasets_to_xlsx(input_names_list, sheet_name_list, writer_name):
    for i in range(len(input_names_list)):
        dataset = get_input_dataset('dataset', i)
        dataset_df = dataset.get_dataframe()
        dataset_df.to_excel(writer, sheet_name= sheet_name_list[i], index = False)
        logger.info("Writing"+ input_names_list[i])
    writer.save()
    
datasets_to_xlsx(input_names, sheet_names, writer)