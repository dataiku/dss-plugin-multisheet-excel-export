#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
Utility functions for conversion to Excel.
Conversion is based on Pandas feature conversion to xlsx.
"""

import logging
import pandas as pd
import os

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format='Multi-Sheet Excel Exporter | %(levelname)s - %(message)s')


def dataframes_to_xlsx(input_dataframes_names, xlsx_abs_path, dataframe_provider):
    """
    Write the input datasets into the same Excel file
    :param input_dataframes_names: List of input dataframe names
    :param xlsx_abs_path: Absolute path of the output Excel file
    :param dataframe_provider: Function to provide the dataframes
    """
    logger.info("Writing output xlsx file ...")
    file_exists = os.path.exists(xlsx_abs_path)
    mode = 'a' if file_exists else 'w'
    writer = pd.ExcelWriter(xlsx_abs_path, engine='openpyxl', mode=mode)
    for name in input_dataframes_names:
        df = dataframe_provider(name)
        logger.info("Writing dataset into excel sheet...")
        df.to_excel(writer, sheet_name=name, index=False)
        logger.info("Finished writing dataset {} into excel sheet.".format(name))
    #writer.save()
    writer.close()
    logger.info("Done writing output xlsx file")

