Dim metrics As Range
Dim bins As Range
Dim minCurrentSTI As Double
Dim maxCurrentSTI As Double
Dim minPreviousSTI As Double
Dim maxPreviousSTI As Double
Dim numberOfBins As Integer
Dim spanCurrentSTI As Integer
Dim spanPreviousSTI As Integer
Dim currentSpan As Long
Dim previousSpan As Long
Dim roundBinTo As Integer

Dim rangeCurrentSTI As Range
Dim rangePreviousSTI As Range
Dim dataRegion As Range
Dim binInputLoop As Boolean

Sub calculateMinAndMax()

' define the two columns as named ranges to avoid repetitive typing
Set rangeCurrentSTI = Range(Range("A2"), Range("A2").End(xlDown))
Set rangePreviousSTI = Range(Range("B2"), Range("B2").End(xlDown))

' define a starting point cell to print out the min and the maxes
Set metrics = Range("A2").End(xlToRight).Offset(0, 3)

' calculate the min and max value for the two columns
minCurrentSTI = Application.WorksheetFunction.Min(rangeCurrentSTI)
minPreviousSTI = Application.WorksheetFunction.Min(rangePreviousSTI)
maxCurrentSTI = Application.WorksheetFunction.Max(rangeCurrentSTI)
maxPreviousSTI = Application.WorksheetFunction.Max(rangePreviousSTI)

metrics.Value = "Minimum Current STI"
metrics.Offset(1, 0).Select
ActiveCell.Value = "Minimum Current STI"
ActiveCell.Offset(0, 1).Value = minCurrentSTI
metrics.Offset(2, 0).Select
ActiveCell.Value = "MaximumCurrent STI"
ActiveCell.Offset(0, 1).Value = maxCurrentSTI
metrics.Offset(3, 0).Select
ActiveCell.Value = "Minimum Previous STI"
ActiveCell.Offset(0, 1).Value = minPreviousSTI
metrics.Offset(4, 0).Select
ActiveCell.Value = "Maximum Previous STI"
ActiveCell.Offset(0, 1).Value = maxPreviousSTI
End Sub

Sub createBins()

Dim binTotal As Long

' declare a few of the variables needed

' set the labels of the different bins
Set bins = metrics.Offset(6, 0)

bins.Select
ActiveCell.Value = "Current STI Bins"
ActiveCell.Offset(0, 1).Value = "Count"

bins.Offset(0, 4).Select
ActiveCell.Value = "Previous STI Bins"
ActiveCell.Offset(0, 1).Value = "Count"

' initially set this condition to false, until the input for the user satisifies the condition
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


Sub generateBinCounts()

bin.Offset(1, 0).Select

Dim binsRange As Range
Set binsRange = Range(ActiveCell, ActiveCell.End(xlDown))

For Each binValue In binsRange
        Count = 0
        low = binValue.Value
        high = binValue.Offset(1, 0).Value
        For Each stiValue In rangeCurrentSTI
            If stiValue.Value



    Next binValue

End Sub


Sub main()

Call calculateMinAndMax
Call createBins
Call generateBinCounts

End Sub