# Cobuild guidance

Workbook behavior:
- The recipe creates one worksheet per `input_dataset` and uses dataset names as worksheet names unless a custom mapping overrides them.
- Set `output_workbook_name` without an extension; the recipe appends `.xlsx` when writing the workbook.
- When `renaming_sheets=true`, each `dataset_to_sheet_mapping` entry has the shape `{"dataset_name":"<input dataset>","sheet_name":"<worksheet name>"}`. Each `dataset_name` must identify one of the selected inputs; inputs without a mapping keep their default worksheet names.
