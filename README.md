# sheets-redundant-info-auto-completer

I made this Apps Script for Google Sheets.  It will iterate over a sheet's contents looking for names or locations.  Using these as primary keys it will then fill in the remaining information using import range function



# Example

Here is an example sheet setup for use:
https://docs.google.com/spreadsheets/d/1tymcrKFLi2SAbdWT8UtGyG30Auav0kPMscI2pQ_9mx0/edit?usp=sharing

Here is the example source sheet it draws from:
https://docs.google.com/spreadsheets/d/1TKCxe3hugvi1d3UoWfbBXNW46S7puqafryx1bv230zI/edit?usp=sharing

# Setup

To configure this for your own needs simply  open the .gs file and paste it into your destination workbook's script editor.
Update each of the following pairs of variables with the appropriate data:

* `hrWorkbookId`, `clientWorkbookId`:  
  Workbook id of the source workbook. i.e. from the example source above: `'1TKCxe3hugvi1d3UoWfbBXNW46S7puqafryx1bv230zI'`

* `hrSheetName`, `clientSheetName`:   
  Name of the sheet with the data in the source workbook.  In the example source above `hrSheetName = 'Personnel'` and `clientSheetName = 'Locations'`.

* `hrGrabbedInfo`, `clientGrabbedInfo`:  
  These arrays denote all of the headers that will be searched and inserted into matching rows.  Check that they are the spelled and written the same on both sheets.

* `hrPrimaryKey`, `clientPrimaryKey`:  
  Names of the primary keys that guide the range insertion.  Check that they are the spelled and written the same on both sheets.  In the example source above `hrPrimaryKey = 'Consultant'` and `clientPrimaryKey = 'Location'`.

* Make sure that you have included this line in your appsscript.json Manifest file:  
  ```"oauthScopes": ["https://www.googleapis.com/auth/spreadsheets"]```
  You will have to click through and "Allow Access" for your spreadsheets to interface with eachother.

# Usage

Once you have completed your setup you can fill in the redundant information by using the "Custom Utilities" dropdown and selecting "Import Consultant Info" or "Import Client Info".

If you are experiencing problems, first check that your setup information is correct.  Check to make sure that your primaryKeys are spelled the same in the source and destination sheets.  Furthermore check that your column headers are spelled the same in the source and destination sheets.

If you follow these instructions you will be able to use this same unaltered script for multiple destination sheets.