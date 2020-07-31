# Documentation Utility
## Disclaimer
Script is provided without any guarantee. It has been tested without adverse effects on the machines I own.

## Overview
Documentaiton or some form of it can be difficult in Power BI. This utility generates three (3) reports (you can run all or a subset of them) that I've found useful for documentation:

| outputDocuments | Description |
| --- | --- |
| DataDescriptions | A table of all the measures and columns in your model with any descriptions you've given them and a visual container ID. |
| VisualContainers | Information on the visual containers in the report (including an ID that matches to DataDescriptions) |
| UnusedObjects | By recursively moving from visual container to source, this is a list of columns and measures that are not necessary to render the report in its current form. |

## OS Requirements
The utlity uses WinForms as a GUI when not all parameters are provided. Consequently, it will only run on Windows.

## Installation

### PowerShell
#### Files Needed
+ Power BI Report Download Utility_0.0.0.9000.ps1

#### Instructions
For the first run on a machine, run as an administrator (it will need to install a module). After that, run in the CLI or right click and select "Run with Powershell".

## Usage
### Run
A module will need to be installed. It happens automatically if you run as administration. Only needs to be done once per machine.

| PowerShell |
| --- |
| `SqlServer' |

### Output
The script will output CSV files to the outputDirectory that correspond to the reports you select.
