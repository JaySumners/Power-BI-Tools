# Custom Power Query Functions

## Overview
A collection of specialized functions that are more efficient or more to-the-point than out-of-the-box functions in Power Query. More information on each function is available in the **Usage** section. Following the `tidyverse` example in R, the dot (`.`) has been replaced by an underscore (`_`) (e.g. `Table_RecordLeftOuterJoin` not `Table.RecordLeftOuterJoin`). 

## Installation
All of the functions are saved in *\*.m* files, which are TXT files ( use the *\*.m* extension to remind me what it is). You have to copy and paste the text into a blank query in order to use the function. With that being said, you can use whatever naming convention you want for the function. For example:

1. Copy the contents of *Table.RecordLeftOuterJoin.m" into a blank query (use the Advanced Editor).
2. Name the query *Table.RecordLeftOuterJoin*. Note: you can name the query anything you want, but this will match the description.
3. Anytime you want to usethe function, open the advanced editor and type `Table.RecordLeftOuterJoin(...)`.

## Usage

### Caveats and Recommendations

Unfortunately, Power Query does not have an "import custom function" option, so these have to be added to each pbix in which you want to use them. Additionally, you have to open the Advanced Editor and enter the function (there are other ways of doing this, but I find this to be the easiest). I'd recommend using these functions to optimize a model rather than during the initial exploration phase.

### Current Available Functions

Function | Article(s) | Description
---- | ---- | ----
`Table_RecordLeftOuterJoin()` | [Mimicking Python Dictonaries in Power Query: Part 1](https://www.linkedin.com/pulse/mimicking-python-dictionaries-power-query-m-how-why-sumners-m-i-a-/) | A left outer join using a record. This process is much faster for a one-to-many or many-to-one join than using the standard ""Merge"".
`Table_RecordReplaceValues()` | [Mimicking Python Dictonaries in Power Query: Part 1](https://www.linkedin.com/pulse/mimicking-python-dictionaries-power-query-m-how-why-sumners-m-i-a-/) | A multi-value replace using an established or manually entered record. NOTE: You cannot replace nulls using a record.

## Upcoming Functions

Function | Article(s) | Description
---- | ---- | ----
`Table_RecordLeftOuterJoinMultiKey()` | [Mimicking Python Dictionaries in Power Query: Part 2](https://www.linkedin.com/pulse/mimicking-python-dictionaries-power-query-m-multi-key-jay/) | A multi-key version of `Table,RecordLeftOuterJoin()` or a more flexible version of it. 
