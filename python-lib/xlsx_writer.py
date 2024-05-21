#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
Utility functions for conversion to Excel.
Conversion is based on Pandas feature conversion to xlsx.
"""

import logging

from openpyxl.styles import Alignment, Font, PatternFill, Side
from openpyxl.styles.borders import Border
from openpyxl.styles.colors import WHITE
import pandas as pd

FONT = "Calibri"
SIZE = 11
DATAIKU_TEAL = "FF2AB1AC"

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format='Multi-Sheet Excel Exporter | %(levelname)s - %(message)s')

def format_worksheet(worksheet):
    font = Font(name=FONT, size=SIZE, color=WHITE, bold=True)
    fill = PatternFill("solid", fgColor=DATAIKU_TEAL)

    no_border_side = Side(border_style=None)
    border = Border(left=no_border_side, right=no_border_side, top=no_border_side, bottom=no_border_side)

    alignment = Alignment(vertical='bottom')

    # border styles do not work on full row selection
    header_row = worksheet.row_dimensions[0]
    header_row.font = font
    header_row.fill = fill
    header_row.border = border
    header_row.alignment = alignment

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

        worksheet = writer.sheets[name] # TODO: maybe put getter
        format_worksheet(worksheet)

        logger.info("Finished writing dataset {} into excel sheet.".format(name))
    writer.save()
    logger.info("Done writing output xlsx file")
