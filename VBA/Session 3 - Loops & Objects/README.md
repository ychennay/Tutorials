# Lesson Plan For Session 3

## Primitives
-	Integer
-	Double
-	Boolean 
-	String
-	`Dim`
-	**Exercise**: save a number to a variable and then print the value out in a cell
-	**Exercise**: Find 25th and 75th %ile of total sales value and save it into variables
-	Dim pct75 as Double
-	 `Pct75 = Application.WorksheetFunction.Percentile_Exc(Range(Range("j2"), Range("j2").End(xlDown)), 0.75)`

## 	Objects
-	Range
-	“Set”
-	Why doesn’t below work? Because you didn’t use “Set”
-	Dim m As Object
-	m = Range("a5")
-	Range("n1").Value = m
-	Exercise: set a range object = to A2 to l2
-	Dim rng As Object
-	Set rng = Range(Range("a1"), Range("a1").End(xlToRight)).Offset(1, 0)
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
-	Go from a2 to bottom of table and find rows -  (RANGE).ROWS.COUNT