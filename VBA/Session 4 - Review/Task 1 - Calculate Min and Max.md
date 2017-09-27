# Calculate Minimum and Maximum STI

## Hints

Here's the list of variables you should declare:

| **Primitives** | **Objects** |
| ------------- | ------------- |
| **minSTI**: You should store the value for the minimum STI here, so you can use it later on.  | **currentSTI**: This is a `Range` object that you should declare and `Set` so that you don't have to constantly type in `Range(Range("B2"), Range("B2").End(xlDown))`. This is also useful because later on if you want to use a `For Each` loop.  |
| **maxSTI**: You should store the value for the max STI here, so you can use it later on.  | **metricsCell**: This is a `Range` object that refers to a cell where you'll begin typing in your metrics (`MIN`, `MAX`. It is optional, but useful to have so you don't have to type lots of stuff.  |