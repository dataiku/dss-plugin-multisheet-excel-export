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
