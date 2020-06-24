Current Version: 1.0.0

# Plugin information

This plugin aims at giving the ability for the user to export multi-sheet excel file.

# Prerequisites

The installation setup for this plugin follows the standard Dataiku code environment creqtion procedure. This plugin requires the [openpyxl](https://openpyxl.readthedocs.io/en/stable/) python module.

# How it works

Once the plugin is successfully installed, select the datasets that you want to export as one excel file. Then run the Multi-Sheet Excel Export recipe from the flow. It will create a folder in your flow containing the output `.xls` file. Each sheet of this file contains one dataset and is named after this dataset.
 