You first need to define all the different variables you'll need when completing the tasks.

First, you'll want to have variables that store the `MIN` and `MAX` of the column `STI Range`:

```
Dim minCurrentSTI As Double
Dim maxCurrentSTI As Double
```

You'll also want to have a `Range` object for the column itself so you don't have to keep typing out `Range(Range(...))`:

```
Dim rangeCurrentSTI As Range
```

You will also need to ask the user for two inputs: 1) `numberOfBins`, and 2) `roundBinTo`, which tells our macro to round the bins to the nearest `X` value, where `X` is the number provided by the user (usually `1000`, or `5000`, or `2000`, a nice round number).

```

Dim numberOfBins As Integer
Dim roundBinTo As Integer
```

You will now need to define variables to determine the size of each bin (the `spanCurrentSTI`). You will also want to declare a variable to store the `currentSpan` when you iterate through the data to check where it belongs.

```
Dim currentSpan As Long
Dim spanCurrentSTI As Integer
```

For convenience, I also define a few other `Range` objects to avoid having to type a lot of redundant code:

The `dataRegion` variable refers to the entire dataset (`A2` to `C1000`)

```
Dim dataRegion As Range
```

The `metrics` variable will be where I begin printing out the `MIN` and the `MAX` value for the `Current STI` column.

```
Dim metrics As Range
```

The `bins` variable will be used to quickly reference a cell where I begin printing the bins.

```
Dim bins As Range
```

Finally, the `binsRange` variable is used to define the range of the bins. We'll need this to create the chart. 

```
Dim binsRange As Range
```
# Calculate Min and Max
This first subroutine simply calculates the minimum and maximum of the `Current STI` column, and prints the results out to the right of the data. 

```
Sub calculateMinAndMax()

' define the two columns as named ranges to avoid repetitive typing
Set rangeCurrentSTI = Range(Range("A2"), Range("A2").End(xlDown))

' define a starting point cell to print out the min and the maxes
Set metrics = Range("A2").End(xlToRight).Offset(0, 3)

' calculate the min and max value for the two columns
minCurrentSTI = Application.WorksheetFunction.Min(rangeCurrentSTI)
maxCurrentSTI = Application.WorksheetFunction.Max(rangeCurrentSTI)

metrics.Value = "Minimum Current STI"
metrics.Offset(1, 0).Select
ActiveCell.Value = "Minimum Current STI"
ActiveCell.Offset(0, 1).Value = minCurrentSTI
metrics.Offset(2, 0).Select
ActiveCell.Value = "MaximumCurrent STI"
ActiveCell.Offset(0, 1).Value = maxCurrentSTI

End Sub
```

# Generate Bins
This next subroutine is more complicated. You need to create bins for the histogram you will create.

```
Sub createBins()

Dim binTotal As Long

' declare a few of the variables needed

' set the labels of the different bins 6 rows underneath the metrics section
Set bins = metrics.Offset(6, 0)

bins.Select
ActiveCell.Value = "Current STI Bins"
ActiveCell.Offset(0, 1).Value = "Count"

' initially set these condition to false, until the input for the user satisifies input
binInputLoop = True
roundToInputLoop = True

Do While binInputLoop
    
    ' ask the user for how many bins they'd like
    numberOfBins = InputBox("Please enter the number of bins you'd like your histogram to have.", "Enter Bin Size")
    
    ' if the number of bins is less than 5, then continue looping
    If numberOfBins < 5 Then
        MsgBox ("The number of bins you selected is too small to be meaningful! Pick a number greater than 5.")

    ' if the number of bins is really large, then also continue looping
    ElseIf numberOfBins > maxCurrentSTI / 5 Then
        MsgBox ("The number of bins you selected is too large!")
    
    ' if neither of those things are true, then exit the loop with the inputted value of numberOfBins
    Else
        binInputLoop = False
    End If

Loop

Do While roundToInputLoop
    
    roundBinTo = InputBox("Please enter the number to round to.")
    If roundBinTo < 0 Then
        MsgBox ("The number to round to is too small")
    Else
        roundToInputLoop = False
    End If

Loop

' calculate the intervals between each bin of a histogram
currentSpan = Round(maxCurrentSTI - minCurrentSTI, 0) / numberOfBins

bins.Offset(1, 0).Select
binTotal = 0

Do While binTotal < currentSpan + maxCurrentSTI
    
    ActiveCell.Value = binTotal
    binTotal = binTotal + currentSpan
    binTotal = Round(binTotal / roundBinTo) * roundBinTo
    ActiveCell.Offset(1, 0).Select

Loop

End Sub

```

# Generate Bin Counts
This subroutine then iterates through each of the values in the `Current STI` column to find the count in each bin. 


```
Sub generateBinCounts()
 
bins.Offset(1, 0).Select
 
Set binsRange = Range(ActiveCell, ActiveCell.End(xlDown))
 
ActiveCell.Offset(0, 1).Select
 
For Each binValue In binsRange
 
        Count = 0
        low = binValue.Value
        high = binValue.Offset(1, 0).Value
        For Each stiValue In rangeCurrentSTI.Cells
            If stiValue.Value > low And stiValue.Value < high Then
                Count = Count + 1
            End If
 
            Next stiValue
 
        ActiveCell.Value = Count
        ActiveCell.Offset(1, 0).Select
 
    Next binValue
 
End Sub
```
# Generate Histogram
Finally, this subroutine will generate the actual chart:

```

Sub generateHistogram()
 
    Dim countsRange As Range
    Dim co As ChartObject
    Dim ct As Chart
    Dim sc1 As SeriesCollection
    Dim ser1 As Series
 
    bins.Offset(1, 1).Select
 
    Set countsRange = Range(ActiveCell, ActiveCell.End(xlDown))
    countsRange.Select
    binsRange.Select
   
    Set co = ActiveSheet.ChartObjects.Add(Range("J2").Left, Range("J2").Top, 600, 600)
    co.Name = "Distribution of STI"
 
    Set ct = co.Chart
   
    With ct
 
        .HasLegend = True
        .HasTitle = True
        .ChartTitle.Text = "Distribution of STI"
        .ChartGroups(1).Overlap = 0
        .ChartGroups(1).GapWidth = 0
 
        Set sc1 = .SeriesCollection
        Set ser1 = sc1.NewSeries
        ser1.Name = "Current STI"
        ser1.XValues = binsRange
        ser1.Values = countsRange
        ser1.ChartType = xlColumnClustered
       
    End With
 
 
 
End Sub
 
```


You need this sub in order to run all the above together:

```
Sub main()

Call calculateMinAndMax
Call createBins
Call generateBinCounts
Call generateHistogram

End Sub
```