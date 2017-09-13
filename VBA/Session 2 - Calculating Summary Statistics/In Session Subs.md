```
Sub CalcAvgSalesPrice()
    'Make a new column called Average Sale Price
    Set SalePriceCell = Range("A1").End(xlToRight).Offset(0, 1)
    SalePriceCell.Value = "Sale Price"
    Range("A1").End(xlToRight).Offset(1, -1).Select
    
    ' Calculate Average Sales Price
    Set Temp = Range(ActiveCell, ActiveCell.End(xlDown))
    Temp.Offset(0, 1).Formula = "= J2 / H2"
    
End Sub
```

```
Sub CalcSummaryStats()

    ' Start at first cell in the top left corner and move all the way to the right
    Range("A1").End(xlToRight).Select
    
    ' Set the Total Sales Column as column J
    Set TotalSalesColumn = Range("J2", Range("J2").End(xlDown))
    
    'Go to cell J2, move to the right, and then move 4 columns to the right
    Range("J2").End(xlToRight).Offset(0, 4).Select
    ActiveCell.Value = "Total Sales"
    
    Set TotSales = ActiveCell.Offset(0, 1)
    TotSales.Value = Application.WorksheetFunction.Sum(TotalSalesColumn)
    
    ActiveCell.Offset(1, 0).Value = "Average Sales"
    Set AvgSales = ActiveCell.Offset(1, 1)
    AvgSales.Value = Application.WorksheetFunction.Average(TotalSalesColumn)
    
    ActiveCell.Offset(2, 0).Value = "Min Sales"
    Set MinSales = ActiveCell.Offset(2, 1)
    MinSales.Value = Application.WorksheetFunction.Min(TotalSalesColumn)
    
    ActiveCell.Offset(3, 0).Value = "Max Sales"
    Set MaxSales = ActiveCell.Offset(3, 1)
    MaxSales.Value = Application.WorksheetFunction.Max(TotalSalesColumn)
    
    ActiveCell.Offset(4, 0).Value = "StDev Sales"
    Set MaxSales = ActiveCell.Offset(4, 1)
    MaxSales.Value = Application.WorksheetFunction.StDev_S(TotalSalesColumn)
    
    'Find the Unit Sales Column
    Set UnitSalesColumn = TotalSalesColumn.Offset(0, 2)
    
    ActiveCell.Offset(6, 0).Value = "Average Unit Sale"
    Set AvgUnitSale = ActiveCell.Offset(6, 1)
    AvgUnitSale.Value = Application.WorksheetFunction.Average(UnitSalesColumn)
    
    ActiveCell.Offset(7, 0).Value = "StDev Unit Sale"
    Set StdDevUnitSale = ActiveCell.Offset(7, 1)
    StdDevUnitSale.Value = Application.WorksheetFunction.StDev_S(UnitSalesColumn)
    
    MsgBox ("Summary statistics have been calculated!")
    
End Sub
```

```
Sub FindHighValueTransactions()

    'Select the first value in the Sale Price column
    Range("L2").Select
    
    ' Until the loop reaches an empty cell, it continues
    Do While ActiveCell.Value <> ""
    
       ' If it's high value, highlight it green
        If ActiveCell.Value > 50 Then
            ActiveCell.EntireRow.Interior.ColorIndex = 43
        
        ' Low value highlight it red
        ElseIf ActiveCell.Value < 10 Then
            ActiveCell.EntireRow.Interior.ColorIndex = 22
        End If
        
        ActiveCell.Offset(1, 0).Select
    
    Loop
    
End Sub
```

```
Sub AllTogetherNow()

    Worksheets("Consolidate Macro").Activate

    Call CalcAvgSalesPrice
    Call CalcSummaryStats
    Call FindHighValueTransactions
    
    Worksheets("July").Activate

    Call CalcAvgSalesPrice
    Call CalcSummaryStats
    Call FindHighValueTransactions

End Sub
```