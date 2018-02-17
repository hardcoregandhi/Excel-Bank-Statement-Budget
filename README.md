# Excel-Bank-Statement-Budget
This VBA Script takes .CSV files exported from online banking apps and creates a small budget table. The Script does not include validation checks (minus checking the report location is clear) as to make it incredibly simple to understand, customise and alter.

A Spreadsheet is also included which already contains the script and (with Macro's enabled) will run with ```Ctrl + b```.
You can paste your statement into this sheet, or having this sheet open will make it accessible to other Excel windows

The Categories are Food, Amazon, Fun, Utility, Takeaway, Subscriptions, Misc.

## Adding a Category to the uneditted script
At Line 41 add:
```
Range("J9").Select
  ActiveCell.Formula = "NEWNAME"
  Range("K9").Select
  ActiveCell.Formula = "=SUMIF(D:D,""NEWNAME"",F:F)"
```   
At Line 90 add:
```
If InStr(ActiveCell.Offset(0, 1).Value, "AS IT APPEARS IN YOUR STATEMENT") Then
  ActiveCell.FormulaR1C1 = "NEWNAME"
End If
```
