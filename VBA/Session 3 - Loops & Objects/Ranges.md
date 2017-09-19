# Ranges

The `Range` object is one of the most common types of Objects that you'll work with in VBA. For example, `ActiveCell` is a Range object.

## Range()

Range() creates a new `Range` object. There's many ways of using `Range()`:

### 1. Range(String)

You can use `Range()` with a `String` address of a cell: `Range("A1")`, for instance.

### 2. Range(String, Range)

You can use `Range()` with a combination of `String` addresses and another `Range` object. 

For example, 

```
Range("A2", Range("A2").End(xlDown))
```

The first argument to `Range` is `"A2"`, which is a `String`. The second argument is `Range("A2").End(xlDown)`, which returns a `Range` object (since `.End()` is a method that returns `Range`s).

Under the surface, VBA actually turns that `String` into a `Range` object. So it actually will convert `"A2"` to `Range("A2")` first, and then plug it in:

`Range(Range("A2"), Range("A2").End(xlDown))`

### 3. Range(Range, Range)

You can also combine two `Range` objects to create another `Range`:

`Range(ActiveCell, ActiveCell.End(xlDown))`

Both `ActiveCell` and `ActiveCell.End(xlDown)` are `Range` objects.

### How you cannot use `Range()`

You cannot, however, use `Range()` with only one `Range` argument, like this:

`Range(ActiveCell)`

`Range` always requires at least two `Range` objects. Range looks for either **one** `String` representing a single cell location, or two `Range` objects. 