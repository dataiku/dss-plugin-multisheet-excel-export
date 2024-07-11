#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
Utility functions for conversion to Excel.
Conversion is based on Pandas feature conversion to xlsx.
"""

import logging
import math
from typing import Tuple

from openpyxl.styles import Alignment, Font, PatternFill, Side
from openpyxl.styles.borders import Border
from openpyxl.styles.colors import WHITE
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd

DATAIKU_TEAL = "FF2AB1AC"
LETTER_WIDTH = 1.20  # Approximative letter width to scale column width
MAX_LENGTH_TO_SHOW = 45  # Limit copied from DSS native excel exporter

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="Multi-Sheet Excel Exporter | %(levelname)s - %(message)s")


def style_header(worksheet: Worksheet, font_name: str = "Calibri", font_size: int = 11, font_color: str = WHITE, background_color: str = DATAIKU_TEAL, bold: bool = True):
    """
    Style header of the worksheet
    """

    if worksheet.min_column < 1:
        logger.warn(f"No header row for worksheet {worksheet}. Styling skipped.")
        return

    font = Font(name=font_name, size=font_size, color=font_color, bold=bold)
    fill = PatternFill("solid", fgColor=background_color)

    no_border_side = Side(border_style=None)
    border = Border(left=no_border_side, right=no_border_side, top=no_border_side, bottom=no_border_side)

    alignment = Alignment(vertical="bottom", horizontal="center")

    for header_cell in worksheet[1]:
        header_cell.font = font
        header_cell.fill = fill
        header_cell.border = border
        header_cell.alignment = alignment


def get_column_width(column: Tuple):
    """
    Find optimum column width based on content and header length
    Based on the computations of DSS native excel output formatter
    """

    header = column[0]
    length_header = len(str(header.value))

    sum_length_cells = 0
    max_length_cells = 0
    for cell in column:
        length_cell = len(str(cell.value))
        max_length_cells = max(max_length_cells, length_cell)
        sum_length_cells += length_cell

    # Computations from ExcelOutputFormatter.java ExcelOutputFormatter.footer
    average_length_cell = math.ceil(sum_length_cells / (len(column) + 1))
    max_length_cells = min(max_length_cells, MAX_LENGTH_TO_SHOW)

    if max_length_cells > 2 * average_length_cell:  # if max length much bigger than average
        length_to_show = int((max_length_cells + average_length_cell) / 2)
    else:
        length_to_show = max_length_cells

    length_to_show = max(length_to_show, length_header)

    return length_to_show * LETTER_WIDTH


def auto_size_column_width(worksheet: Worksheet):
    """
    Resize columns based on the length of the header text
    """
    if worksheet.min_column < 1:
        logger.warn(f"No header row for worksheet {worksheet}. Column auto-size skipped.")
        return

    dimension_holder = DimensionHolder(worksheet=worksheet)

    column_indexes = range(worksheet.min_column, worksheet.max_column + 1)
    for index_column, column in zip(column_indexes, worksheet.iter_cols()):

        column_width = get_column_width(column)
        dimension_holder[get_column_letter(index_column)] = ColumnDimension(worksheet, min=index_column, max=index_column, width=column_width)
    worksheet.column_dimensions = dimension_holder


def dataframes_to_xlsx(input_dataframes_names, xlsx_abs_path, dataframe_provider):
    """
    Write the input datasets into same excel into the folder
    :param input_datasets_names:
    :param writer:
    :return:
    """
    logger.info("Writing output xlsx file ...")
    writer = pd.ExcelWriter(xlsx_abs_path, engine="openpyxl")

    for name in input_dataframes_names:
        df = dataframe_provider(name)

        logger.info("Writing dataset into excel sheet...")
        df.to_excel(writer, sheet_name=name, index=False, encoding="utf-8")

        worksheet = writer.sheets.get(name)

        if worksheet is None:
            logger.warn(f"No worksheet for dataset {name}. Written but styling skipped.")
            continue

        logger.info(f"Styling excel sheet...")
        style_header(worksheet)
        auto_size_column_width(worksheet)

        logger.info("Finished writing dataset {} into excel sheet.".format(name))

    writer.save()
    logger.info("Done writing output xlsx file")
