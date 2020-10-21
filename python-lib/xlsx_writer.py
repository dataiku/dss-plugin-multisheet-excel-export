#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
Utility functions for conversion to Excel.
Conversion is based on Pandas feature conversion to xlsx.
"""

import logging
import pandas as pd

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format='Multi-Sheet Excel Exporter | %(levelname)s - %(message)s')


def dataframes_to_xlsx(input_dataframes, xlsx_abs_path):
    """
    Write the input datasets into same excel into the folder
    :param input_datasets_names:
    :param writer:
    :return:
    """
    writer = pd.ExcelWriter(xlsx_abs_path, engine='openpyxl')
    for name in input_dataframes:
        df = input_dataframes[name]
        logger.info("Writing dataset into excel sheet...")
        df.to_excel(writer, sheet_name=name, index=False, encoding='utf-8')
        logger.info("Finished writing dataset {} into excel sheet.".format(name))
    writer.save()
    print("Finished full xls write")
