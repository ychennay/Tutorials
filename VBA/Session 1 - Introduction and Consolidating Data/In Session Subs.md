```
Sub consolidate_data()
    'activates the second worksheet, selects the data and copies it
    Worksheets("Task 2 Min and Max").Activate
    Range("a2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    'activates the consolidated tab and pastes the values after the data
    Worksheets("Task 1 Consolidate Macro").Activate
    Range("a2").Select
    Selection.End(xlDown).Select
    'tells excel to go down exactly one row from last contingous data point
    Selection.Offset(1, 0).Select
    'pastes values and formats
    Selection.PasteSpecial xlPasteValuesAndNumberFormats
End Sub
```

```
Sub SumTotalSalesValue()
    Worksheets("consolidate macro").Activate
    'selects entire column of sales volume and stores it in "Selection"
    Range("j2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    'sets the cell below the last data point in column j to the sum of the column j
    ' this is equivalent to starting at J2 and then pressind "Ctrl Down" and then down again on arrow key and then setting equal to the sum of column J
    Range("j2").End(xlDown).Offset(1, 0).Value = Application.WorksheetFunction.Sum(Selection)
End Sub
```