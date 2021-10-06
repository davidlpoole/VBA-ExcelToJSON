# Export-Excel-to-JSON

## VBA script to export a table from Excel to JSON


## Usage Instructions:

Clone the file from this repo ("Excel to JSON.xlsm").
This file contains:
- an example data table,
- the VBA script,
- a button to run the script.

Replace the example data with your own, however many columns and
rows as you require. Then click the button and the file will be 
output to the same folder as the "Excel to JSON.xlsm" file.
(Could also remove the button and run the script either via:
- the Developer Ribbon > Macros > exportJSON
- a keyboard shortcut (Developer > Macros > Options...)
- a button on the Quick Access Toolbar
- etc...

Filename, folder, etc can be changed within the script.


## Alternative method

Copy the code only version ("ExcelToJSON.bas") into your your own
spreadsheet (as a VBA module), make sure your spreadsheet is saved 
where you want the output to go. (Also it will error if the file 
is not saved at all... TODO: handle error)
