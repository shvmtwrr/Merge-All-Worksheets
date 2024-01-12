An explanation of the Google Apps Script "Merge-All-Worksheets":

Here's what the script does step by step:

1. `getActiveSpreadsheet()`: This function retrieves the currently active Google Sheets spreadsheet.

2. `getSheetByName('Compilation')`: It attempts to find the sheet named "Compilation" within the spreadsheet. If it doesn't exist, it creates a new sheet with that name.

3. `getSheets()`: This function gets all the sheets (tabs) in the spreadsheet, including the "Compilation" sheet.

4. `clear()`: It clears the contents of the "Compilation" sheet to ensure you start with an empty sheet.

5. The script then enters a loop to process each sheet in the spreadsheet, excluding the "Compilation" sheet itself.

6. `getDataRange()`: This function retrieves the data range of the current sheet, which includes all cells with data.

7. `getValues()`: This retrieves the values (text or numbers) in the data range of the current sheet.

8. `getRichTextValues()`: This retrieves the rich text formatting (including fonts, colors, and styles) of the cells in the data range of the current sheet.

9. The script copies both values and rich text formatting from the current sheet to the "Compilation" sheet using `setValues()` and `setRichTextValues()` functions.

The end result is that the "Compilation" sheet will contain a copy of the values and rich text formatting from all sheets in the spreadsheet, except for the "Compilation" sheet itself. This allows you to compile data and formatting from multiple sheets into a single "Compilation" sheet within the same spreadsheet.
