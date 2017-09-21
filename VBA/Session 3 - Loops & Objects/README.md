# Homework for Session 3

1. **For only the October tab**, calculate the median, mean, 2nd percentile,  98th percentileof all the job titles' `Median` column, place the results in column `G`.

2. **For only the October tab**, calculate the spread (max - min value) of each row and place it as a new column in Column `E`.

3. **For only the October tab**, highlight any job titles that are in the bottom 2% or top 2% of all jobs based on spread (ie. have the highest and lowest spreads).

4. **DIFFICULT, but doable**: Write a macro that searches all 12 worksheets to find the maximum value  and Job Title Code for the `25th Percentile` column in the entire file. Have it display the result in a Message Box after it is done searching. **HINT**: you will need to declare a `Double` variable first, and using an `If` statement.

5. **A NICE CHALLENGE**: Write a macro that iterates through all worksheets in this file and finds the maximum Median value and its job title, displaying it as a Message Box. Be careful:
                -the position of the `Median` column changes in certain worksheets! You'll have to program a way from your Macro to dynamically recognize which column it is on.
                -Certain cells in the `Median` column also contain Strings or empty values. This is typical of what you might get in a real life data set. You have to find a way to handle this without your macro breaking!
 
# Lesson Plan For Session 3 (what was covered)

## Primitives
-	`Integer`
-	`Double`
-	`Boolean` 
-	`String`
-	`Dim`
-	**Exercise**: save a number to a variable and then print the value out in a cell
-	**Exercise**: Find 25th and 75th %ile of total sales value and save it into variables
	```
    Dim pct75 as Double
    Pct75 = Application.WorksheetFunction.Percentile_Exc(Range(Range("j2"), Range("j2").End(xlDown)), 0.75)
    ```
## 	Objects
-	Range
-	`Set`
-	Why doesn’t below work? Because you didn’t use “Set”

    ```
    Dim m As Object
    m = Range("a5")
    Range("n1").Value = m
    ```
-	Exercise: set a range object = to A2 to l2
    
    ```
    Dim rng As Object
    Set rng = Range(Range("a1"), Range("a1").End(xlToRight)).Offset(1, 0)
    ```

##	Loops
###	While Loops
- 	Do while (cond) loop
###	For Loops
-   For I = 1 to 6
-   Next i
### For Each
-   For each rng in myCol
-   Next rng
-   Create a loop that checks if the Total Sales Value is above the 75th or below the 25th, if above or below, shade the selection green/ red 
-   Only shade the range in question, not the entire row
-	Use the rng from above 
-	Go from a2 to bottom of table and find rows -  `(RANGE).ROWS.COUNT`