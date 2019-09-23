# Power BI Tools

## Overview
This repository is a store of tools-in-development, primarily in R or Python, for documenting or improving Microsoft Power BI models (e.g. scripts to find unused objects in your model).

**Note: These scripts will only work on Windows**

## Installation
In general, you can download the script and run it in your favorite IDE or from the command line. Each tool will have specific instructions with any special requirements.

Scripts are available in folders for each tool.

## Usage
### Current Tools
These tools are currently in the repository and available for use.

Tools | Purpose | Languages (Version)
--- | --- | ---
*Unused DAX Objects* | Returns objects in you loaded Power BI model that are not necessary to any visualized calculation | R(`3.6.1`), Python(`3.7.4`)

### Future Tools
+ DAX dependencies network graph (R/Python)
+ DAX calculation flowchart (R/Python)
+ DAX object definitions (R/Python)
+ DAX coding/performance flags (R/Python)
+ Power Query query dependencies network graph (R/Python)
+ Power Query query flowchart with transformations (R/Python)
+ Power Query `Get Sources` to list source files (R/Python)
+ Power Query coding/performance flags (R/Python)
+ Model memory consumption with improvements (R/Power BI)
+ Shiny App and/or Power BI pbix to visualize this data in one place
