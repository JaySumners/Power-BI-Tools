# Unused DAX Objects
## Disclaimer
Script is provided without any guarantee. It has been tested without adverse effects on the machines I own.

## Overview
Includes a WinForm GUI, but can also be used on the CLI.
This WinForms utility downloads a Power BI Report on the Service as a PDF or PPTX.

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
| `MicrosoftPowerBIMgmt' |

### Output
The script will output a a PDF or PPTX to the output directory of the Report with the Pages you selected.
