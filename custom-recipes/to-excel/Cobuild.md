# Cobuild guidance

Use this recipe to export multiple input datasets into one Excel workbook, with one worksheet per dataset.

Roles:
- `input_dataset`: required input datasets to export. This role accepts multiple datasets.
- `folder`: required output managed folder where the workbook is written.

Core configuration:
- Set `output_workbook_name` to the workbook name without the `.xlsx` extension.
- Set `export_conditional_formatting=true` only when conditional formatting should be preserved in the exported sheets.
- Keep `renaming_sheets=false` to use dataset names as worksheet names.
- If custom worksheet names are needed, set `renaming_sheets=true` and provide `dataset_to_sheet_mapping` entries. Each entry contains `dataset_name` and `sheet_name`; `dataset_name` must be one of the selected input datasets.

Output behavior:
- The recipe writes `output_workbook_name + ".xlsx"` to the output managed folder.
