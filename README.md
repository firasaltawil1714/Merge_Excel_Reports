# Merge Excel Reports

A Python script that merges all Excel files in a folder into a single workbook with one header row. Built for recurring reports, such as monthly sales files, that share the same columns.

## Features

- Merges every `.xlsx` file in the `reports/` folder into one file
- Keeps a single header row at the top of the result
- Reads files in a fixed (alphabetical) order, so results are consistent between runs
- Skips files whose headers differ from the first file and reports them, instead of mixing mismatched columns
- Ignores Excel temporary files (such as `~$jan.xlsx`) created when a report is open
- Shows a clear message if the folder is missing or contains no Excel files
- Prints a summary of how many files were merged

## Example

A folder of monthly reports:

```
reports/
├── 01_jan.xlsx
├── 02_feb.xlsx
└── 03_mar.xlsx
```

is merged into one `merged_report.xlsx`, with a single header row at the top.

Example output:

```
Merged 3 of 3 files into merged_report.xlsx
```

If a file has different column names, it is skipped and reported:

```
Skipped (different headers): 03_mar.xlsx
Merged 2 of 3 files into merged_report.xlsx
```

## Requirements

- Python 3.8 or newer
- openpyxl

```bash
pip install openpyxl
```

## Usage

1. Put your Excel files in a folder named `reports` next to the script.
2. Run:

```bash
python Merge_Excel_Reports.py
```

3. Find the merged file, `merged_report.xlsx`, in the same folder as the script.

## Project Structure

```
.
├── Merge_Excel_Reports.py
├── reports/               # put your .xlsx files here
└── merged_report.xlsx     # created automatically
```

## Notes and Limitations

- Only `.xlsx` files are supported.
- Only the first (active) sheet of each file is merged.
- All files should have the same header row (same column names in the same order). Files that don't match the first file are skipped.
- An existing `merged_report.xlsx` is replaced on each run.
- Close the output file in Excel before running the script again.

## Libraries Used

`openpyxl`, `pathlib`

## Author

Firas, Arabic-English technical translator and Python automation learner.
GitHub: https://github.com/firasaltawil1714
