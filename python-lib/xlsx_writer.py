#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
Utility functions for conversion to Excel.
Conversion is based on Pandas feature conversion to xlsx.
"""

import logging
import pandas as pd
from openpyxl import load_workbook
import re

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format='Multi-Sheet Excel Exporter | %(levelname)s - %(message)s')

# input_dataframe_names = names of the tables being inserted into template
# df_map = mapping dataframe from datasets to excel
#


def dataframes_to_xlsx(input_dataframes_names, xlsx_abs_path, dataframe_provider):
    """
    Write the input datasets into same excel into the folder
    :param input_datasets_names:
    :param writer:
    :return:
    """
    logger.info("Writing output xlsx file ...")
    writer = pd.ExcelWriter(xlsx_abs_path, engine='openpyxl')
    for name in input_dataframes_names:
        df = dataframe_provider(name)
        logger.info("Writing dataset into excel sheet...")
        df.to_excel(writer, sheet_name=name, index=False, encoding='utf-8')
        logger.info("Finished writing dataset {} into excel sheet.".format(name))
    writer.save()
    logger.info("Done writing output xlsx file")


def dataframes_to_xlsx_template(input_dataframes_names, mapping, tmp_file_path, dataframe_provider):
    """
    Write the input datasets into same excel into the folder
    :param input_datasets_names:
    :param writer:
    :return:
    """
    logger.info("Writing output xlsx file ...")
    logger.info("Load excel template into writer ...")
    loaded_excel_workbook = load_workbook(filename = tmp_file_path)
    writer = pd.ExcelWriter(tmp_file_path, engine='openpyxl')
    writer.excel_workbook = loaded_excel_workbook
    writer.sheets = dict((sheetname.title, sheetname) for sheetname in loaded_excel_workbook.worksheets) # copy existing sheets into writer

    for name in input_dataframes_names:
        df = dataframe_provider(name)

        # get named ranges
        named_range = mapping[name]
        defined_name_range = loaded_excel_workbook.defined_names[named_range]
        named_range_value = defined_name_range.attr_text

        # get the excel sheet name
        sheet_name = re.findall('(^.*)?!', named_range_value)[0].replace("'", "")
        logger.info("sheet name: ", sheet_name)

        # get the cell ranges from the named range (e.g. $A$20:$I$1048576)
        row_col_vals = re.findall('!(.*)', defined_name_range.attr_text)[0]
        logger.info("row/column values: ", row_col_vals)

        # find start column value (e.g. A = 0, B = 1)
        start_col = re.findall(r'\$(.*?)\$', row_col_vals)[0].lower()
        start_col_num = ord(start_col) - 97
        logger.info("start column number: ", start_col_num)

        # find start row value (e.g. 1, 2, 3)
        start_row_str = re.findall(r'\!(.*?)\:', named_range_value)[0]
        start_row_num = int(re.findall(r'([^$]*$)', start_row_str)[0])-1
        logger.info("start row number: ", start_row_num)

        # write datasets to template in temp file path
        logger.info("Writing dataset into excel sheet...")
        df.to_excel(writer, sheet_name, startrow = start_row_num, startcol = start_col_num, index=False)
        logger.info("Finished writing dataset {} into excel sheet.".format(name))

    writer.save()
    logger.info("Done writing output xlsx file")
