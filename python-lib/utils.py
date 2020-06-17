"""
Utility functions for conversion to Excel.
Conversion is based on Pandas feature conversion to xlsx.
"""

import dataiku
from dataiku.customrecipe import *
import logging
import pandas as pd

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format='Multi-Sheet Excel Exporter | %(levelname)s - %(message)s')


def datasets_to_xlsx(input_datasets_names, xlsx_abs_path):
    """
    Write the input datasets into same excel into the folder
    :param input_datasets_names:
    :param writer:
    :return:
    """
    writer = pd.ExcelWriter(xlsx_abs_path, engine='openpyxl')
    for name in input_datasets_names:
        logger.info("Writing dataset {} into excel sheet...".format(name))
        dataset = dataiku.Dataset(name)
        df = dataset.get_dataframe()
        df.to_excel(writer, sheet_name=name, index=False)
        logger.info("Finished writing dataset {} into excel sheet.".format(name))
    writer.save()
