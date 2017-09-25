# Session 4: Review

Your task is to ultimately write a macro that automatically generates histograms for a particular column. 

**Everyone should be able to complete**:

The macro `calculateMinAndMax`, which is used to automatically find the min and max of a particular column of data. 

**Some people should be able to complete**:

The macro `createBins`, which is used to generate the range of histogram bins used to create the final histogram chart. 

**A few people should be able to complete**:

The macro `generateBinCounts`, which is used to find out how many incumbent data points fall into each bin range.

# Extra Review

Look at the code below and answer the following questions:

1. Summarize in a sentence or two what this macro accomplishes.
2. List all the `Objects` and `primitives` inside this macro.
3. List all `methods` that are called inside this macro.
4. What happens if `Application.DisplayAlerts = True` and `Application.DisplayAlerts = False` are removed from the macro? Test this on an Excel file with multiple worksheets.
5. 

```
Sub NameOfSub()

    Dim myWorksheet As Worksheet
    Dim totalCount As Integer
    Dim deleteCount As Integer
    totalCount = 0
    deleteCount = 0

    For Each myWorksheet In ThisWorkbook.Worksheets
        If myWorksheet.Name <> ThisWorkbook.ActiveSheet.Name Then
        Application.DisplayAlerts = False
        myWorksheet.Delete
        Application.DisplayAlerts = True
        deleteCount = deleteCount + 1
        End If

    totalCount = totalCount + 1
    Next myWorksheet

    MsgBox("Total: " & totalCount & " Deleted: " & deleteCount)

End Sub
```
