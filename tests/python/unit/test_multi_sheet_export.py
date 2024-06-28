from xlsx_writer import datasets_to_xlsx

import os
from openpyxl import Workbook
import pandas as pd
import tempfile


def build_worksheet(headers, data):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in data:
        ws.append(row)
    return ws


def test_datasets_to_xlsx():
    output_file_name = 'sample_test.xlsx'
    tmp_output_dir = tempfile.TemporaryDirectory(dir='.')
    output_file = os.path.join(tmp_output_dir.name, output_file_name)

    df1 = build_worksheet(['dfId', 'gender', 'birthdate'], [[1, 'M', '1953/10/5'], [2, 'L', '1053/12/6']])
    df2 = build_worksheet(['dfId', 'gender'], [[2, 'M']])
    df3 = build_worksheet(['dfId'], [[3],[4],[5],[6],[7]])
    tables = {'df1': df1, 'df2': df2, 'df3': df3}
    
    def worksheet_provider(name):
        return tables[name]

    datasets_to_xlsx(['df1', 'df2', 'df3'], output_file, worksheet_provider)

    df1_out = pd.read_excel(output_file, sheet_name="df1")
    df2_out = pd.read_excel(output_file, sheet_name="df2")
    df3_out = pd.read_excel(output_file, sheet_name="df3")
    os.remove(output_file)

    # len on a panda dataframe returns number of rows without the column names
    assert tables['df1'].max_row-1 == len(df1_out)
    assert tables['df2'].max_row-1 == len(df2_out)
    assert tables['df3'].max_row-1 == len(df3_out)

