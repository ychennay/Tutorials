
The below `Sub()` executes all the other subroutines at the same time.
```
Sub RunAll()

    Worksheets("Consolidate Macro").Activate
    Call AverageUnitPrice
    Call CalcSummaryStats2
    Call CalcMinandMaxStats
    Call Challenge
    
    Worksheets("July").Activate
    Call AverageUnitPrice
    Call CalcSummaryStats2
    Call CalcMinandMaxStats
    Call Challenge
    
    MsgBox ("Called All")
    
End Sub

```

```
Sub AverageUnitPrice()

    Set SalesPriceCell = Range("A1").End(xlToRight).Offset(0, 1)
    SalesPriceCell.Value = "Sale Price"
    Range("A1").End(xlToRight).Offset(1, -1).Select
    Set Temp = Range(ActiveCell, ActiveCell.End(xlDown))
    Temp.Offset(0, 1).Formula = "=J2/H2"
    
    
End Sub

```

```
Sub CalcSummaryStats2()

    Range("A1").End(xlToRight).Offset(0, 4).Select
    Set TotalSalesColumn = Range("J2", Range("J2").End(xlDown))
    Range("J2").End(xlToRight).Offset(0, 4).Select
    ActiveCell.Value = "Total Sales for Month"
    Set TotalSales = ActiveCell.Offset(1, 0)
    TotalSales.Value = Application.WorksheetFunction.Sum(TotalSalesColumn)
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = "Average Sales"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = Application.WorksheetFunction.Average(TotalSalesColumn)
    
End Sub
```

```
Sub CalcMinandMaxStats()

    Range("A2").End(xlToRight).Offset(0, 6).Select
    Set UnitPriceColumn = Range("G2", Range("G2").End(xlDown))
    Range("G2").End(xlToRight).Offset(0, 6).Select
    ActiveCell.Value = "Max Unit Price"
    Set UnitPrice = ActiveCell.Offset(1, 0)
    UnitPrice.Value = Application.WorksheetFunction.Max(UnitPriceColumn)
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = "Min Unit Price"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = Application.WorksheetFunction.Min(UnitPriceColumn)
    Range("A2").End(xlToRight).Offset(0, 10).Select
    ActiveCell.Value = "Standard Deviation of Discount %"
    ActiveCell.Offset(1, 0).Select
    Set DiscountColumn = UnitPriceColumn.Offset(0, 2)
    ActiveCell.Value = Application.WorksheetFunction.StDev(DiscountColumn)
    
    MsgBox ("Donzo!")
      
End Sub
```


This subroutine calculates the 20th and 80th percentiles, and highlights them if they fall outside those ranges. Notice that the cell where the value is stored in is a Range object called `Eighty` and `Twenty`: 
```
Sub Challenge()

    Range("G1").Select
    ActiveCell.Offset(1, 0).Select
    Set PriceColumn = Range("G2", Range("G2").End(xlDown))
    ActiveCell.End(xlDown).Offset(3, 0).Value = "Percentile"
    ActiveCell.End(xlDown).Offset(5, 0).Value = Application.WorksheetFunction.Percentile(PriceColumn, 0.8)
    Set Eighty = ActiveCell.End(xlDown).Offset(5, 0)
    ActiveCell.End(xlDown).Offset(7, 0).Value = Application.WorksheetFunction.Percentile(PriceColumn, 0.2)
    Set Twenty = ActiveCell.End(xlDown).Offset(7, 0)
    
    Range("G2").Select
    Do While ActiveCell.Value <> ""
    If ActiveCell.Value > Eighty Then
    ActiveCell.EntireRow.Interior.ColorIndex = 53
    ElseIf ActiveCell.Value < Twenty Then
    ActiveCell.EntireRow.Interior.ColorIndex = 20
    End If
    ActiveCell.Offset(1, 0).Select
    Loop
    
End Sub
```