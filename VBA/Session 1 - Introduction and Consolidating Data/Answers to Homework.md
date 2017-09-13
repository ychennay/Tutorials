```
Sub MinVal()

    Range("I27").Value = Application.WorksheetFunction.Min(Range("J2:J25"))
 
End Sub
```

```
Sub MinVal2()

     Range("I27").Value = Application.WorksheetFunction.Min(Range("J2", Range("J2").End(xlDown)))
     
End Sub
```

```
Sub MaxVal()

    Range("I30").Value = Application.WorksheetFunction.Max(Range("J2:J25"))
    
End Sub
```

```
Sub MaxVal2()

    Range("I30").Value = Application.WorksheetFunction.Max(Range("J2", Range("J2").End(xlDown)))
    
End Sub
```

```
Sub AverageVal()

    Range("N2").Value = Application.WorksheetFunction.Sum(Range("H2:H25"))
    
End Sub
```

```
Sub SumRange()

    Worksheets("Task 3 Copy Offset").Activate
    Range("D3").Select
    Selection.Copy
    Range("D3:D1000").Select
    Range("D3:D1000").Formula = Range("D3").Formula
    Range("D1002").Value = Application.WorksheetFunction.Sum(Range("D1", Range("D1").End(xlDown)))
     
End Sub

```