# Multisheet excel export Plugin

This plugin converts several DSS datasets to one multi-sheet excel (`.xlsx`) file containing one sheet per input dataset.
More information in the [Documentation](https://www.dataiku.com/product/plugins/multisheet-excel-export/)

# Prerequisites

The installation setup for this plugin follows the standard DSS code environment creation procedure.
This plugin relies on the [openpyxl](https://openpyxl.readthedocs.io/en/stable/) Python module.

# How it works

Once the plugin is successfully installed, select the datasets that you want to export as one excel file. 
Then run the Multi-Sheet Excel Export recipe from the flow. 
It will create a folder in your flow containing the output `.xlsx` file. Each sheet of this file contains one dataset and is named after this dataset.
 
## Running tests

In order to run the tests contained in `python-test\`, launch the following command from the plugin root directory: 
`PYTHONPATH=$PYTHONPATH:/path/to/python-lib pytest`


### Licence

Copyright 2020-2022 Dataiku SAS

This plugin is distributed under the Apache License version 2.0
