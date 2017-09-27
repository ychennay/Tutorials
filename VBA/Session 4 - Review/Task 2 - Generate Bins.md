# Generate Bins

## Hints

You should declare some variables to help you keep track of values

| **Primitives** | **Objects** |
| ------------- | ------------- |
| **numberOfBins**: You will have to ask the user how many bins to crete for your histogram. Since this is a whole number, you can save this as an `Integer`. | **bins**: An optional `Range` object to help you mark the location where you will begin printing your bins.  |
| **roundBinTo**: You will also have to ask your user to what number they want to round their bins to (usually a nice clean even number like `100` or `1000` or `500`) |  |

General workflow for this macro:

1. Ask your user for how many bins they'd like to have (use `InputBox("Text you want displayed in your input box")`).
2. Ask your user to what value they'd like to round their bins to.
3. Navigate to the section of your worksheet where you want to begin printing the bins 
4. Create a header label called `"Current STI Bins"`
5. Create a variable to store your running (current) bin value. You'll need this variable in order to know when you should stop creating bins.
6. Below that header, begin counting off your bins from your `MIN` STI value to your `MAX` STI value. You'll almost certainly need to use a `For` loop, a `Do While` loop, or a `For Each` loop.
7. Stop after your bin value is greater than the `MAX` STI value.

## Checklist for Success

1. Does your `Sub()` generate bins that are inclusive of `MIN` and `MAX` STI values (otherwise, the largest and smallest data points won't be shown in your histogram!)?

2. **Advanced**: What happens if your user doesn't input a number and instead inputs a character? Or inputs a negative number? Or inputs a decimal for the number of bins? What can you add to your code to make sure your macro doesn't break and quit?