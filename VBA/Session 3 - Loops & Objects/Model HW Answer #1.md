```
' HOMEWORK 2 STARTS HERE

Sub AvgSalesPrice()

' SalesPriceCell = name of cell itself; sets as variable?

    Set SalesPriceCell = Range("A1").End(xlToRight).Offset(0, 1)
    SalesPriceCell.Value = "Sale Price"

' Needs the -1 otherwise it'll go allllll the way down

    Range("A1").End(xlToRight).Offset(1, -1).Select
    Set Temp = Range(ActiveCell, ActiveCell.End(xlDown))

' But we want L so...

    Temp.Offset(0, 1).Formula = "=J2/H2"

End Sub
```

```
Sub SummaryStats()

    Range("A1").End(xlToRight).Offset(0, 4).Select

' Give column J a name - declare a range

    Set TotalSalesColumn = Range("J2", Range("J2").End(xlDown))
    Range("J2").End(xlToRight).Offset(0, 4).Select
    ActiveCell.Value = "Total Sales Per Month"

    Set TotalSales = ActiveCell.Offset(0, 1)
    TotalSales.Value = Application.WorksheetFunction.Sum(TotalSalesColumn)

' Active Cell is still P2 because that's the last thing we selected - not Q2

    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = "Average Sales"
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Application.WorksheetFunction.Average(TotalSalesColumn)

' Row for min

    ActiveCell.Offset(1, -1).Select
    ActiveCell.Value = "Minimum Sale Value"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Application.WorksheetFunction.Min(TotalSalesColumn)

' Row for max

    ActiveCell.Offset(1, -1).Select
    ActiveCell.Value = "Maximum Sale Value"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Application.WorksheetFunction.Max(TotalSalesColumn)

' Row for standard deviation

    ActiveCell.Offset(1, -1).Select
    ActiveCell.Value = "Standard Deviation"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Application.WorksheetFunction.StDev(TotalSalesColumn)

' Rows for 25th and 75th percentiles

    ActiveCell.Offset(1, -1).Select
    ActiveCell.Value = "25th percentile"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Application.WorksheetFunction.Percentile(TotalSalesColumn, 0.25)
    
    ActiveCell.Offset(1, -1).Select
    ActiveCell.Value = "75th percentile"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Application.WorksheetFunction.Percentile(TotalSalesColumn, 0.75)
    
End Sub
```

```
' Highlight transactions under the 20th percentile and over the 80th percentile

Sub HighlightPercentileValues()

    ' Select the first value in the Sale Price column
    Range("L2").Select
  
    ' Until the loop reaches an empty cell, it continues
    Do While ActiveCell.Value <> ""
    
    ' Made a temporary placeholder to evaluate the percentile of the active cell
    ' Note: do NOT put "Set" before TempPercentile because it will freak out - "Set" is used for simple data types (integer, string, etc)
    
        TempPercentile = Application.WorksheetFunction.PercentRank(Range("L2", Range("L2").End(xlDown)), ActiveCell.Value)
        
       ' Evaluate the placeholder to see if it's above 0.8
       If TempPercentile > 0.8 Then
            ActiveCell.EntireRow.Interior.ColorIndex = 36
        
        ' Evaluate the placeholder to see if it's less than 0.2
       ElseIf TempPercentile < 0.2 Then
            ActiveCell.EntireRow.Interior.ColorIndex = 20
       End If
        
       ' Move on to the next cell!
        ActiveCell.Offset(1, 0).Select
    
   Loop
  

End Sub
```

```
' Overarching sub making summary stats & highlights for both worksheets

Sub SummaryStatsBothWorksheets()

    Worksheets("Consolidate Macro").Activate
    Call AvgSalesPrice
    Call SummaryStats
    Call HighlightPercentileValues
    
    Worksheets("July").Activate
    Call AvgSalesPrice
    Call SummaryStats
    Call HighlightPercentileValues
    
    MsgBox ("All Done :)")

End Sub
```