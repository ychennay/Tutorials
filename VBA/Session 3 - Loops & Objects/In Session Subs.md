```
Sub dowhileloopPractice()
    Dim r As Integer
    r = 1
    Do While (r < 10)
        Range("z" & r).Value = r
        r = r + 1
    Loop
End Sub
```

```
Sub forLoopPractice()
    For i = 1 To 10
        Range("aa" & i).Value = i
    Next i
End Sub
```

```
Sub forEachLoop()
    Dim entry As Object
    Set entry = Range(Range("i2"), Range("i2").End(xlDown))
    
    Dim distance As Integer
    distance = Application.Match("Total Sales Value", Range(Range("a1"), Range("a1").End(xlToRight)), 0)
    
    Dim total As Double
    
    
    For Each pct In entry
        If pct.Value = 0.15 Then
            total = total + 1
        End If
    Next pct
    Range("n1").Value = total
End Sub
```

```
Sub highlightHighsAndLows()
    Dim TSV As Object
    Set TSV = Range(Range("j2"), Range("j2").End(xlDown))
    
    Dim myRange As Object
    Set myRange = Range(Range("a2"), Range("a2").End(xlToRight))
    
    Dim pct75 As Double
    Dim pct25 As Double
    
    pct75 = Application.WorksheetFunction.Percentile_Exc(TSV, 0.75)
    pct25 = Application.WorksheetFunction.Percentile_Exc(TSV, 0.25)
    
    Dim count As Integer
    count = 2
    Do While (Range("a" & (count)).Value <> "")
        If Range("j" & count).Value > pct75 Then
            myRange.Interior.ColorIndex = 43
        ElseIf (Range("j" & count).Value) < pct25 Then
            myRange.Interior.ColorIndex = 22
        End If
        
        Set myRange = myRange.Offset(1, 0)
        count = count + 1
    Loop
End Sub
```

```
Sub test()
    'Range(Range("j2"), Range("j2").End(xlDown))
    
    Dim rng As Object
    Set rng = Range(Range("a1"), Range("a1").End(xlToRight)).Offset(1, 0)
    rng.Clear
    
    'Range("n1").Value = Application.WorksheetFunction.Percentile_Exc(Range(Range("j2"), Range("j2").End(xlDown)), 0.75)
    Dim m As Object
    Set m = Range("a5:a25")
    
    'for each loop
    For Each rng In m
        Range("n5").Value = m.Rows.count
    Next rng
    
    Range("L2").Select
    
    'While loop
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
    
    Dim i As Integer
    i = 1
    
    For i = 1 To 25
        Range("a" & i + 5).Value = Range("a" & i + 5).Value * 2
    Next i
    
End Sub
```