#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
Utility functions for conversion to Excel.
Conversion is based on Pandas feature conversion to xlsx.
"""

import logging
import math
from typing import Tuple, List, Dict
from copy import copy

from openpyxl.styles import Alignment, Font, PatternFill, Side
from openpyxl.styles.borders import Border
from openpyxl.styles.colors import WHITE
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook

DATAIKU_TEAL = "FF2AB1AC"
LETTER_WIDTH = 1.20 # Approximative letter width to scale column width
MAX_LENGTH_TO_SHOW = 45 # Limit copied from DSS native excel exporter
EXCEL_MAX_LEN_SHEET_NAME = 31

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format='Multi-Sheet Excel Exporter | %(levelname)s - %(message)s')

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
    
    if max_length_cells > 2 * average_length_cell: # if max length much bigger than average
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
        dimension_holder[get_column_letter(index_column)] = ColumnDimension(worksheet, 
                                                           min=index_column, 
                                                           max=index_column,
                                                           width=column_width)
    worksheet.column_dimensions = dimension_holder




# code inspired from https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/worksheet/copier.html
def copy_sheet_to_workbook(source_sheet: Worksheet, target_workbook: Workbook) -> Worksheet:
    """
    Copy the source worksheet as a new worksheet in the target workbook
    :param source_sheet: the source sheet
    :param target_workbook: the workbook used to store the new sheet
    :return: a reference to the created sheet inside the workbook
    """
    logger.info(f"Copying sheet {source_sheet.title} to target workbook")
    target_sheet = target_workbook.create_sheet(source_sheet.title)
    for row in source_sheet:
        for cell in row:
            new_cell = target_sheet.cell(row=cell.row, column=cell.column, value=cell.value)
            new_cell.data_type = cell.data_type
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.alignment = copy(cell.alignment)

    return target_sheet

def rename_too_long_dataset_names(input_dataset_names: List[str]) -> Dict[str, str]:
    """
    Excel allows for only maximum 30 chars in the sheet names, so if some DS have more than 30 chars :
        - truncate the name to 28 chars
        - Add an index from 00 to 99 at the end in case of overlap
    :param input_dataset_names: the list of dataset names to remap
    :returns: a Dict[str, str] mapping the DS names with the sheet names
    """

    return_map = {}
    index_rename = -1
    renaming_length = EXCEL_MAX_LEN_SHEET_NAME - 2
    for name in input_dataset_names:
        if len(name) > EXCEL_MAX_LEN_SHEET_NAME:
            index_rename += 1
            rename = f"{name[0:renaming_length]}{index_rename:02d}"
            # Almost impossible case : a DS already has this name
            while rename in input_dataset_names:
                index_rename += 1
                rename = f"{name[0:renaming_length]}{index_rename:02d}"

            logger.info(f"Dataset {name} with a too long name will be stored as sheet {rename}")
            return_map[name] = rename
        else:
            return_map[name] = name

    return return_map


def datasets_to_xlsx(input_dataset_names, xlsx_abs_path, worksheet_provider):
    """
    Write the input datasets into same excel into the folder
    :param input_dataset_names: the list of dataset to put in a single excel file, using one sheet (excel tab) per dataset
    :param xlsx_abs_path: the temporary path where to write the final excel file
    :param dataset_provider: a lambda used to get the dataset
    """

    logger.info(f"Building output excel file {xlsx_abs_path}")
    # The final workbook where all dataset sheets will be written
    workbook = Workbook()
    # remove the default sheet created
    workbook.remove(workbook.active)

    renaming_map = rename_too_long_dataset_names(input_dataset_names)

    for name in input_dataset_names:
        dataset_worksheet = worksheet_provider(name)
        if dataset_worksheet is None:
            continue

        if name in renaming_map:
            dataset_worksheet.title = renaming_map[name]
        else:
            # should never happen
            logger.warn(f"Failed to find a name for the workshhet {name}")
            dataset_worksheet.title = name
        
        target_sheet = copy_sheet_to_workbook(dataset_worksheet, workbook)
        
        logger.info(f"Styling excel sheet {target_sheet.title} in target workbook")
        auto_size_column_width(target_sheet)

        logger.info(f"Finished writing dataset {name} into excel sheet.")
        
    workbook.save(xlsx_abs_path)
    logger.info("Done writing output xlsx file")
