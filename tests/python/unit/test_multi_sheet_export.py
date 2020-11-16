from xlsx_writer import dataframes_to_xlsx

import os
import pandas as pd
import tempfile


def test_datasets_to_xlsx():

    output_file_name = 'sample_test.xlsx'
    tmp_output_dir = tempfile.TemporaryDirectory(dir='.')
    output_file = os.path.join(tmp_output_dir.name, output_file_name)

    df1 = pd.DataFrame({'dfId': [1], 'gender': ['M'], 'birthdate': ['1953/10/5']})
    df2 = pd.DataFrame({'dfId': [2], 'gender': ['M']})
    df3 = pd.DataFrame({'dfId': [3]})
    tables = {'df1': df1, 'df2': df2, 'df3': df3}

    def dataframe_provider(name):
        return tables[name]

    dataframes_to_xlsx(['df1', 'df2', 'df3'], output_file, dataframe_provider)

    df = pd.read_excel(output_file, sheet_name="df1")
    os.remove(output_file)

    assert len(tables["df1"]) == len(df)

