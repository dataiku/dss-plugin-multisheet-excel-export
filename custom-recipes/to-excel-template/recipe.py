#!/usr/bin/python
# -*- coding: utf-8 -*-
"""
Custom recipe for Multisheet Export to Existing Excel Template
"""

import logging

from pathvalidate import ValidationError, validate_filename

import dataiku
from dataiku.customrecipe import get_input_names_for_role
from dataiku.customrecipe import get_output_names_for_role
from dataiku.customrecipe import get_recipe_config

from cache_utils import CustomTmpFile
from xlsx_writer import dataframes_to_xlsx_template

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format='Multi-Sheet Excel Export to Existing Excel Template | %(levelname)s - %(message)s')

### Get User Input Values ###
# read in input datasets
input_datasets_ids = get_input_names_for_role('input_dataset2')
if len(input_datasets_ids) == 0:
    logger.warning("Received no input datasets ids. input_datasets_ids={}".format(input_datasets_ids))
# make a list of input datasets
input_datasets_names = [name.split('.')[-1] for name in input_datasets_ids]
if len(input_datasets_names) == 0:
    logger.warning("Received no input datasets names. input_datasets_ids={}, input_datasets_names={}".format(
        input_datasets_ids, input_datasets_names))


# get input folder containing template
input_folder_with_template_id = get_input_names_for_role('input_folder_with_template')
input_folder_with_template_name = input_folder_with_template_id[0]
input_folder_template = dataiku.Folder(input_folder_with_template_name)

# get path of excel template in folder
excel_template_name = input_folder_template.list_paths_in_partition()[0][1:]

# retrieve the output folder_id
output_folder_id = get_output_names_for_role('folder')
logger.info("Retrieved the following folder ids: {}".format(output_folder_id))
output_folder_name = output_folder_id[0]
logger.info("Received the following output folder name: {}".format(output_folder_name))
output_folder = dataiku.Folder(output_folder_name)

# set up output file and folder name
input_config = get_recipe_config()
workbook_name = input_config.get('output_workbook_name', None)

if workbook_name is None:
    logger.warning("Received input received recipe config: {}".format(input_config))
    raise ValueError('Could not read the workbook name.')

output_file_name = '{}.xlsx'.format(workbook_name)

try:
    validate_filename(output_file_name)
except ValidationError as e:
    raise ValueError(f"{e}\n")

# set up named range mapping
mapping = input_config.get('mapping')
for dataset in input_datasets_names:
    if dataset in mapping:
        continue
    else:
        logger.warning("Received input received recipe config: {}".format(input_config))
        raise ValueError('Could not find the named range mapping for dataset: {}'.format(dataset))

### Start Work ###

# Create Temporary file
tmp_file_helper = CustomTmpFile()
tmp_file_path = tmp_file_helper.get_temporary_cache_file(output_file_name)
logger.info("Intend to write the output xls file to the following location: {}".format(tmp_file_path))

# Save template in temp path
with input_folder_template.get_download_stream(excel_template_name) as stream:
    data = stream.read()
    stream.close()

with open(tmp_file_path, "wb") as binary_file:
    # Write bytes to file
    binary_file.write(data)


# Iterate through the input datasets, and insert into appropriate sheet and location in template (stored in temp file path)
dataframes_to_xlsx_template(input_datasets_names, mapping, tmp_file_path, lambda name: dataiku.Dataset(name).get_dataframe())

# Save file to output folder
with open(tmp_file_path, 'rb', encoding=None) as f:
    output_folder.upload_stream(output_file_name, f)

tmp_file_helper.destroy_cache()

logger.info("Ended recipe processing.")
