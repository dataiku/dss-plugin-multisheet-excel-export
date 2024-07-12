# Changelog

## [Version 2.0.0](https://github.com/dataiku/dss-plugin-multisheet-excel-export/releases/tag/v2.0.0) - Major release - 2024-07
- Important : Column type changed ! From this version, cell types in excel will reflect the storage type in DSS. For example, string column containing only numbers will be exported as text column. If you want a number column in excel, you need to have a integer/float column on DSS
- Export dataset conditional formatting colors (colors the cells, does not export rules)
- Bug fix : can now export dataset with date types

## [Version 1.1.4](https://github.com/dataiku/dss-plugin-multisheet-excel-export/releases/tag/v1.1.4) - Bug release - 2024-06
- Fix numpy issue with DSS 13

## [Version 1.1.3](https://github.com/dataiku/dss-plugin-multisheet-excel-export/releases/tag/v1.1.3) - Minor release - 2024-06
- The plugin will show up in the excel category

## [Version 1.1.2](https://github.com/dataiku/dss-plugin-multisheet-excel-export/releases/tag/v1.1.2) - Feature release - 2024-05
- Style exported worksheet header
- Auto-size columns to fit data

## [Version 1.1.1](https://github.com/dataiku/dss-plugin-multisheet-excel-export/releases/tag/v1.1.0) - Chore release - 2023-08
- Use python library to create temp file instead of a custom cache

## [Version 1.1.0](https://github.com/dataiku/dss-plugin-multisheet-excel-export/releases/tag/v1.1.0) - New hotfix release - 2023-08
- Updated plugin to python 3.7, 3.8, 3.9, 3.10, 3.11

## [Version 1.0.1](https://github.com/dataiku/dss-plugin-multisheet-excel-export/releases/tag/v1.0.1) - New hotfix release - 2022-03
- Changed documentation URL
