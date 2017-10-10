# Session 6: Headcount

***This will be our final VBA session for now. We've really enjoyed working with and teaching everyone! Please let us know your feedback, recommendations, and ideas for future sessions! ***

In this exercise, we will work on using an organization's dataset files to calculate headcount per month, ultimately creating a "calendar of headcounts" per month. These headcounts will populate this (currently blank) table:

![VLookup](/VBA/Images/output.png)

### Dataset:
- two columns (**start date** and **end date**): the `start date` will reflect when a person joins the workforce, and `end date` reflects when the person leaves the workforce

### Workflow and Pseudo-code:

1. Look at data and calculate min and max (non-VBA is okay)
2. Construct the empty output graph (shown above)
    - columns are months
    - rows are years

3. Load `start date` and `termination date` month and year into separate variables
4. There should be a total of **4** variables
    - Denote current year and month with `y` and `m` respectively
    - Use the VBA equivalents of the Excel functions `MONTH()` and `YEAR()` to parse the dates into month and year variables

5. Create a **while loop**: while `y` is less than `termination year`
    - In each iteration of the loop...
    - Add `1` to the year and month index in your output from step 1 (this denotes that in the current `y` and `m` index, the person is currently employed in that month)
    - Use **index match** or **VLOOKUPS**
    - Check if `y` and `m` are equal to `termination year` and `termination month`. If so, exit the loop
    - Check if `m` == 12. If so, set m = 1 (January) and increase `y` by 1 (go to next year)

     By iterating through this loop, you will go through each month and year and add one to the graph, and ultimately arrive at the total headocunt. 

Questions? Email us!