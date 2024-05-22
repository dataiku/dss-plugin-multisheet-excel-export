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
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd

DATAIKU_TEAL = "FF2AB1AC"
LETTER_WIDTH = 1.23 # Approximative letter width to scale column width

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format='Multi-Sheet Excel Exporter | %(levelname)s - %(message)s')

def style_header(worksheet: Worksheet, 
                  font_name: str = "Calibri", 
                  font_size: int = 11, 
                  font_color : str = WHITE, 
                  background_color : str = DATAIKU_TEAL,
                  bold : bool = True
                 ):
    """
    Style header. worksheet must have at least one row
    """
    font = Font(name=font_name, size=font_size, color=font_color, bold=bold)
    fill = PatternFill("solid", fgColor=background_color)

    no_border_side = Side(border_style=None)
    border = Border(left=no_border_side, right=no_border_side, top=no_border_side, bottom=no_border_side)

    alignment = Alignment(vertical='bottom', horizontal='center')

    for header_cell in worksheet[1]:
        header_cell.font = font
        header_cell.fill = fill
        header_cell.border = border
        header_cell.alignment = alignment

def auto_size_column_width(worksheet: Worksheet):
    """
    Resize columns based on the lenght of the header
    worksheet must have at least one row
    """
    dimension_holder = DimensionHolder(worksheet=worksheet)

    for column in range(worksheet.min_column, worksheet.max_column + 1):
        column_letter = get_column_letter(column)
        header_cell = worksheet[f"{column_letter}1"]
        column_target_width = len(header_cell.value) * LETTER_WIDTH

        dimension_holder[column_letter] =  ColumnDimension(worksheet, 
                                                           min=column, 
                                                           max=column,
                                                           width=column_target_width)
    worksheet.column_dimensions = dimension_holder


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

        worksheet = writer.sheets.get(name)

        if worksheet is None:
            logger.warn(f"No worksheet for dataset {name}. Written but styling skipped.")
            continue

        if worksheet.min_column < 1:
            logger.warn(f"No header row for dataset {name}. Written but styling skipped.")
            continue

        logger.info(f"Styling excel sheet...")
        style_header(worksheet)
        auto_size_column_width(worksheet)

        logger.info("Finished writing dataset {} into excel sheet.".format(name))

    writer.save()
    logger.info("Done writing output xlsx file")
