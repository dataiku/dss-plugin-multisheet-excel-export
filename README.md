# Multisheet excel export

This plugin enables the export of multiple DSS datasets to a single multi-sheet excel file (*.xlsx*). For each exported dataset, there will be one sheet in the output file.

## Requirements

The installation setup for this plugin follows the standard Dataiku code environment creation procedure. Through code environment creation at plugin install time, the following packages will be installed

- openpyxl==3.0.6
- pathvalidate==2.3.0

## How To Use

### Export multiple DSS datasets as one multi-sheet Excel file

From your flow, select the datasets you want to merge in a single Excel file,

Click on the **Multisheet Excel Export** icon in the column at the right of the flow or from the recipe menu,

This plugin contains a single recipe that merges the input dataset in a single (multisheet) Excel worksheet,

Choose a name for the folder that will contain the output file in the flow,

Choose a name for the output excel worksheet. (Do not include a file extension in this name, the extension will always be *.xlsx*)

The resulting folder will appear in the flow as shown in the following screenshot,

You can now click on the folder and download the *.xslsx*file.
