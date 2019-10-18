# Power BI on Powershell

## Overview

A set of Powershell scripts to help manage Power BI. Scripts may include data wrangling, but will generally utilize the Power BI Rest API or (preferably) the Power BI Powershell cmdlet.

Note to self: put link to here to setup and basics of Power BI Powershell cmdlet and Power BI Rest API (as well as links to documentation).

## Installation

Download the script and run in Powershell. Depending on your settings, this may be a double-click on the *\*.ps1* file or you may have to open Powershell (or the Powershell ISE) and run the file. Nothing in these should perform any operation that would be detremental to your system, **however, it is good practice to review ANY script before running it on your machine.** Additionally, I may provide Power BI or other files to bring together the output of some scripts.

Cmdlets or modules may be required if you don't already have them downloaded. Normally, the code to install anything required will be commented out at the top of the script.

## Usage

Some scripts may include a full or partial GUI interface and others will operate entirely in the console. 

### Current Available Tools

Tool | Special Permissions | Resources | Description
---- | ---- | ---- | ----
`Get Pro License Users.ps1` | None | | Returns all Power BI Pro licenses assigned within your organization. 

### Future Tools

Tool | Special Permissions | Resources | Description
---- | ---- | ---- | ----
`Get Power BI User Activity.ps1` | Power BI Admin | | Get user activity throught the tenant.
`Get & Combine Power BI Pro Users & Activity.ps1` | Power BI Admin | | Return  a CSV with all Power BI Pro license holders and their most recent activty.
`Remove Broken Gateways.ps1' | Power BI Admin | | Remove broken gateways.
`Visualize Licenses & Activity.pbix` | None | | Visualizes the output of `Get Pro License Users.ps1` and `Get Power BI User Activity.ps1`. 
`Visualize Power BI Activity.pbix` | None | | Visualize the output of `Get Power BI User Activity.ps1`.


## Disclaimer

As with everything on this repo, these scripts are provided as-is and with no guarantee. Please review any script before running on your machine.
