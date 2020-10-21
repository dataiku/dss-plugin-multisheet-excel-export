from unittest import TestCase
from xlsx_writer import dataframes_to_xlsx
from cache_utils import CustomTmpFile
import os
import pandas as pd


class Test(TestCase):

    def setUp(self) -> None:
        self.input_data = './data'
        self.output_folder = './python-test/test_output'
        self.names = []
        self.dataframes = []
        available_csv = os.listdir(self.input_data)
        self.tables = {}
        for csv_name in available_csv:
            csv_path = os.path.join(self.input_data, csv_name)
            csv_parsed_name = csv_name[:-4]
            csv_df = pd.read_csv(csv_path)
            self.names.append(csv_parsed_name)
            self.dataframes.append(csv_df)
            self.tables[csv_parsed_name] = csv_df

    def test_datasets_to_xlsx(self):
        output_file = os.path.join(self.output_folder, 'sample_test.xlsx')
        dataframes_to_xlsx(self.tables, output_file)
        df = pd.read_excel(output_file, sheet_name="customers")
        assert len(self.tables["customers"]) == len(df)
        os.remove(output_file)
