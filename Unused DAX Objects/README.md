**SUPERSEDED BY [Documentation Utility](https://github.com/JaySumners/Power-BI-Tools/tree/master/Documentation%20Utility)**


# Unused DAX Objects
## Overview
After building a model or making significant changes, I tend to end up with a number of measures or columns (or even tables) that are no longer connected any visualization (directly or by calculation). This script recursively searches dependencies and lists those objects that are not necessary to produce the report in its current form.

## Installation

### R
#### Files Needed
+ Unused Objects - v1.0.0.0.R
+ powershell_scripts folder

I've also included an `Unused Dax Objects.Rproj` for ease of use.

#### Instructions
Put all files in the same directory and run as usual.

### Python
#### Files Needed
+ Unused Objects - v1.0.0.0.py
+ powershell_scripts folder

#### Instructions
Put all files in the same directory and run as usual.

## Usage
### Run
You may need to install a few libraries, but nothing out of the ordinary.

| R | Python |
| --- | --- |
| `tidyverse`, `jsonlite` | `tkinter`, `os`, `shlex`, `subprocess`, `re`, `pandas`, `numpy`, `zipfile`, `json` |

### Output
The script will output a CSV with the following columns:

| Table | Object  | Description |
| ---   | ---     | ---         |
| Table where the column or measure lives | The object name in [Name] format | Any description you've added to the object |
