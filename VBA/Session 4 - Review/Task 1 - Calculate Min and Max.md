# Calculate Minimum and Maximum STI

## Hints

Here's the list of variables you should declare:

| **Primitives** | **Objects** |
| ------------- | ------------- |
| **minSTI**: You should store the value for the minimum STI here as a `Double` or `Long`, so you can use it later on. Don't use an `Integer`, since sometimes the maximum values for STI might be potentially larger than the range of numbers an `Integer` primitive variable can store. | **currentSTI**: This is a `Range` object that you should declare and `Set` so that you don't have to constantly type in `Range(Range("B2"), Range("B2").End(xlDown))`. This is also useful because later on if you want to use a `For Each` loop.  |
| **maxSTI**: Same as above. | **metricsCell**: This is a `Range` object that refers to a cell where you'll begin typing in your metrics (`MIN`, `MAX`. It is optional, but useful to have so you don't have to type lots of stuff.  |

The general workflow to create this `Sub()`:

1. Declare your variables and objects
2. Dynamically select from the top of the dataset down and set it to your `Range` object.
3. Calculate the `MIN` and the `MAX` of that range.
4. Navigate to some spot away from your dataset on your worksheet, and print the results there (but also store them in your variables!)

## Checklist for Success

- Does your `Sub()` calculate `MIN` and `MAX` successfully, regardless of how many rows of data there are?
- Does your `Sub()` calculate `MIN` and `MAX` successfully, regardless of how many *columns* of data there are?
- Are your calculated values stored as variables for future use?

