## Homework for next session:

Finish the rest of the macro. Try not to look at the code we wrote in today's session. For both `Consolidate Data` and `July` tab, calculate the `MIN()`, `MAX()`, `STDEV.S` (standard deviation of a sample), and `25th` and `75th` percentiles. 

* Try to write the code without looking at the example code. If you're stuck, look at the `Common VBA Functions` document, and then [Google](www.google.com) or [StackOverflow](www.stackoverflow.com). If you're stuck, then you may refer to the example code in `In Session Subs.md`. Of course, your code will probably look similar since there's only a finite number of ways to write these macros, but the whole point is to practice yourself. Anyone can copy and paste code and hit the run button.
* Your macro should execute on both tabs (remember that you can use `Worksheet("NameOfYourSheet").Activate` to switch to another sheet.)
* Your macro should be able to work regardless of the size of data (number of rows) or number of variables (# of columns). Make sure to use **relative**, as opposed to **absolute** references. 
* You should have a main `Sub()` that calls other subroutines to organize your code more logically. Remember that you can execute other subs by writing

```
Call NameOfSubYouWantToExecute
Call NameOfAnotherSub
```
* Your macro should end by alerting the user that it has completed with `You're Done!` or another message of your choosing using a `Message Box`.

### Challenge:

Look at the example in the sub `FindHighValueTransactions()` in the `In Session Subs.md` file to see how I highlighted a column for high (defined as above `$50`) and low (defined as under `$5`) values. The code uses a `DO WHILE` loop: the loop continues until the statement is false. In the example, it keeps looping until `ActiveCell.Value <> ""`, which can be translated to until the active cell is an empty value. Then, without copying, write your own code to search through the `Average Unit Price` column and highlight any transactions that are under the `20th percentile` and above the `80th percentile`. 