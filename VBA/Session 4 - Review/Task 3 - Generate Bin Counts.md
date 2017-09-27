# Generate Bin Counts

Prior to this `Sub()`, you've calculated the `MIN` and the `MAX` for the `Current STI` column in your dataset. Now, you'll need to find a way to find out how many employees fall into each "bin".

## Hints

| **Primitives** | **Objects** |
| ------------- | ------------- |
| **low**: This is the start of your bin, so for an employee to fall into a particular bin, he/she needs to have a `Current STI` greater than this value. Probably should be a `Double`. | **currentRangeSTI**: An `Range` object that consists of all the values in the `Current STI` column. |
| **high**: This is the end of your bin, so for an employee to fall into a particular bin, he/she needs to have a `Current STI` **lower** than this value. You know that if `ActiveCell` is your `low` variable, then `ActiveCell.Offset(1,0)` is your `high` value. Probably should be a `Double`. | **binRange**: A `Range` object that consists of all the bins. You want to define this as a `Range` so you can loop through each value. |

## Proposed Workflow

You will probably need **two** loops, one within another. You'll need the outer loop to loop through each `bin` value and calculate `low` and `high` values, but you'll also need an inner loop to then iterate (loop) through the `Current STI` column. 

You will definitely need an `If` statement inside the inner loop. If you want to check for multiple conditions, you'd write:
```
If condition1 And condition2 Then

    'Do some stuff if true

End If
```
I bring up multiple conditions since you have to check that a particular employee's `Current STI` is both **greater** than `low` and **less** than `high`.

## Checklist for Success

1. Check the counts that your `Sub()` produces with the official Microsoft `histogram` package. Go to `Data Analysis` in the `Data` tab and select the histogram option when the input box appears. 

2. **Advanced**: what happens if there is a data entry issue and one of the values in the `Current STI` column is actually a string? Or what if it's empty? Incorporate some way of handling these issues that won't break your macro.

3. **Super Advanced**: my proposed workflow is not the most efficient way of generating counts. In fact, it actually quite inefficient. How could you structure the code so that it is more efficient computationally? You don't necessary need to actually write the more efficient code, but simply explain how you'd go about it.