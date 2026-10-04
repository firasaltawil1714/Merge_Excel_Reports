# Merge_Excel_Reports

Combines multiple Excel files with identical column structure into a single merged file.

## Assumption
All input files must have the same header row (same column names, same order). The script does not currently validate this — mismatched headers will be merged incorrectly without warning.

## Example
A folder of monthly sales reports (`jan.xlsx`, `feb.xlsx`, `mar.xlsx`) is merged into one `merged_report.xlsx`, with a single header row at the top.

## Requirements

pip install openpyxl


## How to run

python excel_merger.py
