# Cobuild guidance

- Use an existing managed folder for `folder`. If the user did not specify one, ask them to provide one or create it in DSS; do not invent a folder reference.
- The recipe creates one worksheet per `input_dataset`, in input order, and uses dataset names as worksheet names unless a custom mapping overrides them.
- When `renaming_sheets=true`, each `dataset_to_sheet_mapping` entry has the shape `{"dataset_name":"<input dataset>","sheet_name":"<worksheet name>"}`. Each dataset must be a selected input, each sheet name must be at most 31 characters, and unmapped inputs keep their default worksheet names.
