# Convert Excel to CSV
## Overview
CSV files load significantly faster than sheets from an Excel document. Opening each Excel file and saving the sheet we want can be a tedious process. The PowerShell scripts here automate that process (to a degree). Running the script will allow you to easily extract a sheet from an Excel file and save it as a CSV.

## Installation
No special instructions. The files are self contained and you only need one of them (see the Usage -> Run section).

## Usage
### Run
#### Interactive Script
Run by right clicking on the `.ps1` file and selecitng `Run with PowerShell` or open PowerShell and run like normal (remember the `&` before a string to run it and not just print out the string). The script will give you a seleciton box to pick you Excel file and then one to select the sheet you want to convert.

#### Command Line Script
The command line script is mostly for programmatic calling (e.g. from Python). To get a CSV of `Sheet1` of Excel file `C:\some folder\myExcel.xlsx` with the Excel to CSV script held in `C:/scripts/Excel to CSV (Command Line).ps1`:

Running from PowerShell:
```powershell
& "C:\scripts\Excel to CSV (Command Line).ps1" "C:\some folder\myExcel.xlsx" "Sheet1"
```

### Output
The output will be a CSV file in the format `excelFileName_SheetName.csv`. If you have an Excel file called `myExcel.xlsx` and you want to convert `Sheet1`, then you'll get out `myExcel_Sheet1.csv`. **Note: the csv will save to the same directory as the Excel file.**
