# Table of Contents

<!-- TOC -->

- [Table of Contents](#table-of-contents)
    - [**Formatting**](#formatting)
        - [***Color***](#color)
            - [Color Palette Table](#color-palette-table)
            - [Set the background color of cell `A1` to red](#set-the-background-color-of-cell-a1-to-red)
            - [Set the background color of an entire row of cells to a certain color](#set-the-background-color-of-an-entire-row-of-cells-to-a-certain-color)
        - [**Font**](#font)
            - [Set the font in `A9` to bold](#set-the-font-in-a9-to-bold)
            - [Set the font in `A2` to be regular](#set-the-font-in-a2-to-be-regular)
            - [Set the cell `B4` to be both bold and italic:](#set-the-cell-b4-to-be-both-bold-and-italic)
- [**Variables**](#variables)
    - [**Types of Primitive Variables**](#types-of-primitive-variables)
    - [**Declaring Variables**](#declaring-variables)
            - [Force VBA to explicitly declare variables](#force-vba-to-explicitly-declare-variables)
        - [***Module Level***](#module-level)
            - [Declare a `Project` Level variable](#declare-a-project-level-variable)
            - [Declare a `Module` Level variable](#declare-a-module-level-variable)
            - [Declaring a String (text) variable, assign the variable the value of cell `A1`, and then assign the value of cell `C1` the value of the variable](#declaring-a-string-text-variable-assign-the-variable-the-value-of-cell-a1-and-then-assign-the-value-of-cell-c1-the-value-of-the-variable)
- [**Basic Functions**](#basic-functions)
    - [**Working With New Sheets and Workbooks**](#working-with-new-sheets-and-workbooks)
        - [***Sheets***](#sheets)
            - [Delete the worksheet `Sheet2`](#delete-the-worksheet-sheet2)
            - [Save the name of the active sheet to a variable](#save-the-name-of-the-active-sheet-to-a-variable)
            - [Add a new sheet after the active sheet](#add-a-new-sheet-after-the-active-sheet)
            - [Change sheet orientation to landscape](#change-sheet-orientation-to-landscape)
    - [**Other**](#other)
        - [***Formatting your code***](#formatting-your-code)
            - [Separating your code onto multiple lines](#separating-your-code-onto-multiple-lines)
        - [***Address of cells and ranges***](#address-of-cells-and-ranges)
            - [Find the address of the active cell](#find-the-address-of-the-active-cell)
            - [Get the address of the last cell in the column `A`](#get-the-address-of-the-last-cell-in-the-column-a)
        - [***Clearing and Deleting Stuff***](#clearing-and-deleting-stuff)
            - [Clear (delete) the contents of column `G`](#clear-delete-the-contents-of-column-g)
            - [Clear the contents of all cells](#clear-the-contents-of-all-cells)
            - [Assign a range of cells to a variable](#assign-a-range-of-cells-to-a-variable)
            - [Assigning Value to Cell](#assigning-value-to-cell)
            - [Assign the value `"Yu Chen"` to a range of cells from `A1` to `D2`](#assign-the-value-yu-chen-to-a-range-of-cells-from-a1-to-d2)
            - [Assign a value of `"Yu Chen"` to the variable `MyVariable`, and then assign this variable to cell `B2`](#assign-a-value-of-yu-chen-to-the-variable-myvariable-and-then-assign-this-variable-to-cell-b2)
            - [Assigning Formula to Cell](#assigning-formula-to-cell)
    - [**Selecting Things**](#selecting-things)
        - [***Selecting Workbooks***](#selecting-workbooks)
            - [Activating the current workbook (where the code resides)](#activating-the-current-workbook-where-the-code-resides)
            - [Activate a workbook to the name of `My Macro Book`](#activate-a-workbook-to-the-name-of-my-macro-book)
            - [Activate the 2nd workbook (or specifically, the workbook in index position `2`)](#activate-the-2nd-workbook-or-specifically-the-workbook-in-index-position-2)
        - [***Selecting Cells***](#selecting-cells)
            - [**Select a single cell `A2`**](#select-a-single-cell-a2)
            - [Select a group of cells that are not next to each other](#select-a-group-of-cells-that-are-not-next-to-each-other)
            - [Select the cells from `A2` to `B10`](#select-the-cells-from-a2-to-b10)
            - [Select the last row in column `A` of the dataset, and then moves `6` rows down](#select-the-last-row-in-column-a-of-the-dataset-and-then-moves-6-rows-down)
            - [Select the last row in column `C` of the dataset, and then moves `6` rows down](#select-the-last-row-in-column-c-of-the-dataset-and-then-moves-6-rows-down)
            - [Select the entire region of cells](#select-the-entire-region-of-cells)
            - [Select entire column that `A2` is on (column `A`)](#select-entire-column-that-a2-is-on-column-a)
            - [Select entire row that `A2` is on (row `2`)](#select-entire-row-that-a2-is-on-row-2)
            - [Select from the active cell down to the last row](#select-from-the-active-cell-down-to-the-last-row)
            - [Select fom the active cell to the last column and save it as a `Range` object called `MyRange`](#select-fom-the-active-cell-to-the-last-column-and-save-it-as-a-range-object-called-myrange)
        - [***Selecting Sheets***](#selecting-sheets)
            - [Select a sheet by tab name (`Sheets2`)](#select-a-sheet-by-tab-name-sheets2)
            - [Select the next sheet in your workbook](#select-the-next-sheet-in-your-workbook)
            - [Select the previous sheet in your workbook](#select-the-previous-sheet-in-your-workbook)
        - [***Selecting Worksheets***](#selecting-worksheets)
            - [**Select the `Task 2 Min and Max` tab**](#select-the-task-2-min-and-max-tab)
        - [**Copying / Pasting Things**](#copying--pasting-things)
            - [Copy the value in cell `M3`](#copy-the-value-in-cell-m3)
            - [Copy and paste value from cell `A1` to `B1` all in one line](#copy-and-paste-value-from-cell-a1-to-b1-all-in-one-line)
            - [Assign the cell `M2` the value `10`. Assign the cell `M3` the formula `=SUM(M2,2)`, which should equal `12`. Copy this formula. Paste this formula from `M4` down to `M100`.](#assign-the-cell-m2-the-value-10-assign-the-cell-m3-the-formula-summ22-which-should-equal-12-copy-this-formula-paste-this-formula-from-m4-down-to-m100)
    - [**Formulas**](#formulas)
        - [***General***](#general)
            - [Assigns each cell from `D40` to `F40` the formula found in cell `F28`](#assigns-each-cell-from-d40-to-f40-the-formula-found-in-cell-f28)
            - [Assigns each cell from `D40` to `F40` a particular formula](#assigns-each-cell-from-d40-to-f40-a-particular-formula)
            - [Assign each cell from `D40` to `F40` the formula found in cell `F28`](#assign-each-cell-from-d40-to-f40-the-formula-found-in-cell-f28)
        - [***Formulas with Row Column Notation***](#formulas-with-row-column-notation)
            - [Assign the active cell the formula found in `G2` using Row Column notation](#assign-the-active-cell-the-formula-found-in-g2-using-row-column-notation)
            - [Assign the active cell the formula `= 65 - the cell one to the left of the active cell`](#assign-the-active-cell-the-formula--65---the-cell-one-to-the-left-of-the-active-cell)
    - [**Using Excel Functions**](#using-excel-functions)
        - [***Math Functions***](#math-functions)
            - [Find the AVERAGE value of the range from `A2` to `A26` and place it in cell `A28`](#find-the-average-value-of-the-range-from-a2-to-a26-and-place-it-in-cell-a28)
            - [Find the minimum of column `A`](#find-the-minimum-of-column-a)
            - [Find the max of row `3`](#find-the-max-of-row-3)
            - [Select all the values starting at `A2` down (`CTRL + DOWN`) in Excel, then sums the values](#select-all-the-values-starting-at-a2-down-ctrl--down-in-excel-then-sums-the-values)
            - [Use `VLOOKUP` to look up the cell value in `E2` from the second column of reference table `A1:C20`](#use-vlookup-to-look-up-the-cell-value-in-e2-from-the-second-column-of-reference-table-a1c20)
    - [**User Interaction**](#user-interaction)
        - [***Message Boxes***](#message-boxes)
            - [A simple message box with the text "Learning is kewl! and "OK" Button](#a-simple-message-box-with-the-text-learning-is-kewl-and-ok-button)
            - [A message box with two lines of text on it](#a-message-box-with-two-lines-of-text-on-it)
            - [A message box with the text "Learning is kewl!" and "Breaking News..." as the title](#a-message-box-with-the-text-learning-is-kewl-and-breaking-news-as-the-title)
            - [Read the text from cell `A1` and display it inside the message box](#read-the-text-from-cell-a1-and-display-it-inside-the-message-box)
            - [Ask the user if they would like to proceed or not in a message box](#ask-the-user-if-they-would-like-to-proceed-or-not-in-a-message-box)
        - [***Dialog Boxes***](#dialog-boxes)
            - [Save one tab of a multi-tab workbook to a new file using the `Save As` dialog box](#save-one-tab-of-a-multi-tab-workbook-to-a-new-file-using-the-save-as-dialog-box)
        - [***Input Boxes***](#input-boxes)
            - [Ask a user for their first and last name in an input box and greet them with text in cell `C3`.](#ask-a-user-for-their-first-and-last-name-in-an-input-box-and-greet-them-with-text-in-cell-c3)
            - [Ask user to enter their name, and validate that input was entered using `IF` condition](#ask-user-to-enter-their-name-and-validate-that-input-was-entered-using-if-condition)
            - [Ask user to click a cell and return the address of that cell](#ask-user-to-click-a-cell-and-return-the-address-of-that-cell)
            - [Ask for a decimal number value in the Input Box and place a type check to validate](#ask-for-a-decimal-number-value-in-the-input-box-and-place-a-type-check-to-validate)
    - [**Loops**](#loops)
        - [***Do Loops***](#do-loops)
            - [Start at `A1` and move down until reaching an empty cell](#start-at-a1-and-move-down-until-reaching-an-empty-cell)
            - [Start at `A1` and move down until reaching a cell with a value less than `50`](#start-at-a1-and-move-down-until-reaching-a-cell-with-a-value-less-than-50)
        - [***For Next Loops***](#for-next-loops)
            - [Loop from `A1` down 20 cells](#loop-from-a1-down-20-cells)
            - [Loop from `A1` down `20` cells, with steps of `2`](#loop-from-a1-down-20-cells-with-steps-of-2)
    - [**Logic and Conditions**](#logic-and-conditions)
        - [***If Statements***](#if-statements)
            - [Check if the active cell has a value of `5`](#check-if-the-active-cell-has-a-value-of-5)
            - [Check whether or not the user clicked `Yes` in a message box](#check-whether-or-not-the-user-clicked-yes-in-a-message-box)
        - [Select Case](#select-case)
            - [Check the value of cell `A1`](#check-the-value-of-cell-a1)
- [Debugging](#debugging)

<!-- /TOC -->

## **Formatting**

### ***Color***

#### Color Palette Table

![Colors](/VBA/Images/colors_vba.gif)

#### Set the background color of cell `A1` to red

`Range("A1").Interior.Color = vbRed`

([Home](#table-of-contents))
#### Set the background color of an entire row of cells to a certain color

`Range("A1").EntireRow.Interior.Color = vbRed`

([Home](#table-of-contents))
### **Font**

#### Set the font in `A9` to bold
`Range("A9").Font.Bold = True`

You can also set it this way:

`Range("A9").Font.FontStyle = "Bold"`
([Home](#table-of-contents))
#### Set the font in `A2` to be regular
`Range("A9").Font.Bold = False`

You can also set it this way:

`Range("A9").Font.FontStyle = "Regular"`
([Home](#table-of-contents))
#### Set the cell `B4` to be both bold and italic:

`Range("B4").Font.FontStyle = "Bold italic"`
([Home](#table-of-contents))
# **Variables**

Here are some general naming conventions when declaring and using variables:

- must be less than **255** characters
- must only use letters, numbers, or underscores (**no spaces**!)
- cannot begin with a number

There are different levels of **Scope** for a variable. Scope means what context the variable can be used.

- **Procedure Level**: this variable can only be used inside the subroutine that it has been declared in. They are declared as using `Dim`.
- **Module Level**: this variable can be used in all the subroutines in the current module. You declare the variable the same way as the **Procedure Level**, but place it before all of the `Sub()` declarations. 
- **Project Level**: this can be used anywhere in the current project. You declare this type of variable with the keyword `Public`.

## **Types of Primitive Variables**

| **Type** | **Description** | **Memory Used**
| ------------- | ------------- | ------------- |
|  **Boolean**  | A logical status of either `TRUE` or `FALSE`. This status also corresponds to `0` and `1`.  | 2 bytes |
|  **Integer**  |Any whole number between **-32,768** and **32,767**. Use this for **discrete** counts (counting things that cannot be split into decimals or fractions, such as # of successful attempts, # of employees). You probably shouldn't use this, for example, when counting money. | 2 bytes	|
|  **Long**  |Any whole number between **-2,147,483,648** and **2,147,483,647**. In general, you won't need to use this variable type, unless you are working with extremely large numbers (such as the population of the world), or GDP of countries. This type takes up twice as much space as a regular `Integer`. | 4 bytes |
|  **Single**  |Any number (whole or decimal) between **-3.4e38** and **-1.4e-45** for negative values and **1.40e-45** through **3.40e38** for positive values. Unless you need extreme precision with your calculation (ie. calculating microscope changes in radioactive decay), this is probably the data type you need for decimals  | 4 bytes |
|  **Double**  |Any number (whole or decimal) between **-1.78e308** and **-4.94e-324** for negative values and **4.9.40e-324** through **1.8e308** for positive values.  | 8 bytes |
|  **Currency**  | Used to represent dollar values from **-922,337,203,685,477.5808** to **-922,337,203,685,477.5807**  | 8 bytes |
|  **Date**  | Used to represent dates from **January 1st, 100 CE** to **December 31st, 9999 CE**  | 8 bytes |
|  **String**  | Text used to store names, descriptions, etc.  | 10 bytes in addition to string length |
|  **Variant**  | Any of the above. Because it is the most flexible and can be anything, it also takes up the most space, since VBA has to reserve a lot of room in case it's a large data type. | At least 16 bytes.

([Home](#table-of-contents))
## **Declaring Variables**

#### Force VBA to explicitly declare variables

At the top of your code, before you write your subroutines, write

`Option Explicit`

This means that if you attempt to run the code below:

```
Sub SubC()
myNewVariable = "Hello"
End SubC()
```

You will receive a `Compile error: Variable not defined` error, since the variable `myNewVariable` was not explicitly defined with

`Dim myNewVariable As String`

([Home](#table-of-contents))
### ***Module Level***

#### Declare a `Project` Level variable

`Public myName As String`

Project variables are declared using the keyword `Public`. It is available to all subroutines in the entire project.

#### Declare a `Module` Level variable

`Dim strModuleLevelVariable As String`

This variable is declared just like any other variable, except its value is available for all subroutines that are found inside the module. You declare the variable **before** the sub routines, like this:

```
Dim strModuleLevelVariable As String
strModuleLevelVariable = "Yu Chen"

Sub Sub1()
Range("A3).Value = strModuleLevelVariable
End Sub1()

Sub Sub2()
Range("B5").Value = strModuleLevelVariable
End Sub2()
```

In both cell `B5` and `A3`, the value of `Yu Chen` will be inputted, since both `Sub1()` and `Sub2()` are in the same module.

#### Declaring a String (text) variable, assign the variable the value of cell `A1`, and then assign the value of cell `C1` the value of the variable

```
Dim myNewStringVar As String
Range("A1").Select
myNewStringVar = ActiveCell.Value
Range("C1").Value = myNewStringVar
```

([Home](#table-of-contents))
# **Basic Functions**

## **Working With New Sheets and Workbooks**

### ***Sheets***

#### Delete the worksheet `Sheet2`
`Worksh­eet­s("S­hee­t2").Delete`

#### Save the name of the active sheet to a variable
```
Dim MySheet As String

Sub SomeSub()
MySheet = ActiveSheet.Name
End Sub
```
The reason why the variable `MySheet` is declared as a module variable is because you'll probably use it as you switch between different subroutines.

#### Add a new sheet after the active sheet

`Sheets.Add After:=ActiveSheet`

#### Change sheet orientation to landscape

`ActiveSheet.PageSetup.Orientation = xlLandscape`

## **Other**

### ***Formatting your code***

#### Separating your code onto multiple lines

Sometimes a particular function in VBA will take up a lot of space to write. One way to keep your code neat is to break it up into multiple lines. In order to tell VBA that you'd like to break up  your code into multiple lines, using ` _ ` to separate the code.

Here is an example:
` Temp = Application.WorksheetFunction.VLookup(Range("C2"), Ranage("A4:B200"),2,False)`

can be broken up into

```
Temp = Application.WorksheetFunction.VLookup _ 
    (Range("C2"), Range("A4:B200), 2, False)

```

([Home](#table-of-contents))

### ***Address of cells and ranges***

#### Find the address of the active cell
`ActiveCell.Address`

([Home](#table-of-contents))
#### Get the address of the last cell in the column `A`
`Range("A1").End(xlDown).Address`

([Home](#table-of-contents))
### ***Clearing and Deleting Stuff***

#### Clear (delete) the contents of column `G`
`Range("G:G").ClearContents`

([Home](#table-of-contents))
#### Clear the contents of all cells
`Cells.ClearContents`

([Home](#table-of-contents))
#### Assign a range of cells to a variable
`MyRange = Range("A1:B20")`

([Home](#table-of-contents))
#### Assigning Value to Cell
`Range("M2").Value = 10`

Note here that when you are entering in a number, you do not need to put in quotation marks! If you were to write this instead:

`Range("M2").Value = "10"`

You would get a string value, not a integer (number) value.

([Home](#table-of-contents))
#### Assign the value `"Yu Chen"` to a range of cells from `A1` to `D2`
`Range("A1:D2").Value = "Yu Chen"`

([Home](#table-of-contents))

#### Assign a value of `"Yu Chen"` to the variable `MyVariable`, and then assign this variable to cell `B2`
```
MyVariable = "Yu Chen"
Range("B2").Value = MyVariable
```
([Home](#table-of-contents))

#### Assigning Formula to Cell
This assigns the Excel formula to `M3` (take the value of `M2` and add `2` to it.)

`Range("M3").Formula = =SUM(M2,2)`

([Home](#table-of-contents))
## **Selecting Things**

([Home](#table-of-contents))
### ***Selecting Workbooks***

([Home](#table-of-contents))
#### Activating the current workbook (where the code resides)
`ThisWorkbook.Activate`

([Home](#table-of-contents))
#### Activate a workbook to the name of `My Macro Book`
`Workbooks("My Work Book").Activate`

([Home](#table-of-contents))
#### Activate the 2nd workbook (or specifically, the workbook in index position `2`)
`Workbooks(2).Activate`

([Home](#table-of-contents))
### ***Selecting Cells***

([Home](#table-of-contents))
#### **Select a single cell `A2`**

`Range("A2").Select`

([Home](#table-of-contents))
#### Select a group of cells that are not next to each other
`Range("B2,C8,E9").Select`

`Range("B2,C8,E9:E20").Select`

([Home](#table-of-contents))
#### Select the cells from `A2` to `B10`

`Range("A2", "B10").Select`

([Home](#table-of-contents))
#### Select the last row in column `A` of the dataset, and then moves `6` rows down 
`Range("A1").End(xlDown).Offset(6, 0).Select`

([Home](#table-of-contents))
#### Select the last row in column `C` of the dataset, and then moves `6` rows down
`Range("C1").End(xlDown).Offset(-6, 0).Select`

([Home](#table-of-contents))
#### Select the entire region of cells
This is equivalent to hitting `CTRL + SHIFT + DOWN + RIGHT` on your keyboard:

`ActiveCell.CurrentRegion.Select`

This is also equivalent to the following command:

`Range("A2", Range("A2").End(xlDown).End(xlToRight)).Select`

([Home](#table-of-contents))
#### Select entire column that `A2` is on (column `A`)

`Range("A2").EntireColumn.Select`

([Home](#table-of-contents))
#### Select entire row that `A2` is on (row `2`) 
`Range("A2").EntireRow.Select`

([Home](#table-of-contents))
#### Select from the active cell down to the last row

```
Range(ActiveCell, ActiveCell.End(xlDown))
```

([Home](#table-of-contents))
#### Select fom the active cell to the last column and save it as a `Range` object called `MyRange`

```
Dim MyRange As Range
Set MyRange = Range(ActiveCell, ActiveCell.End(xlDown))
```

([Home](#table-of-contents))
### ***Selecting Sheets***

A `Sheet` and a `Worksheet` are related, but cannot be used interchangeably. A `Sheet` is any Excel sheet, whereas a `Worksheet` is only a regular Excel worksheet. For example, a chart is a `Sheet` but is not a `Worksheet`.

([Home](#table-of-contents))
#### Select a sheet by tab name (`Sheets2`)
`Sheets("Sheets2").Select`

([Home](#table-of-contents))
#### Select the next sheet in your workbook

`ActiveSheet.Next.Select`

([Home](#table-of-contents))
#### Select the previous sheet in your workbook

`ActiveSheet.Previous.Select`

([Home](#table-of-contents))
### ***Selecting Worksheets***

#### **Select the `Task 2 Min and Max` tab**
`Worksheets("Task 2 Min and Max").Activate`

([Home](#table-of-contents))
### **Copying / Pasting Things**

([Home](#table-of-contents))
#### Copy the value in cell `M3`
`Range("M3").Copy`

([Home](#table-of-contents))
#### Copy and paste value from cell `A1` to `B1` all in one line
`Range("A1").Copy Range("B1")`

Note that this pastes all the formatting as well, so if you had a bolded cell in `A1`, you'll also have a bolded cell in `B1`.

([Home](#table-of-contents))
#### Assign the cell `M2` the value `10`. Assign the cell `M3` the formula `=SUM(M2,2)`, which should equal `12`. Copy this formula. Paste this formula from `M4` down to `M100`.

```
Range("M2").Value = 10 
Range("M3").Formula = "=SUM(M2,2)"
Range("M3").Copy 
Range("M4:D100").PasteSpecial
```
([Home](#table-of-contents))

## **Formulas**

### ***General***

#### Assigns each cell from `D40` to `F40` the formula found in cell `F28` 
`Range("D40:F40").Formula = Range("F28").Formula`

([Home](#table-of-contents))
#### Assigns each cell from `D40` to `F40` a particular formula

`Range("D40:F40").Formula = "=65 - A2"`

Please note that when you do this, the formula is `F28` is interpreted **relatively**. For example, if your formula is `= 65 - A2`, the first row will have `= 65 - A2`, but the sec ond row will have `= 65 - A3` and the third row will have `= 65 - A4`.

([Home](#table-of-contents))
#### Assign each cell from `D40` to `F40` the formula found in cell `F28`
`Range("D40:F40").Formula = Range("F28").Formula`

### ***Formulas with Row Column Notation***

#### Assign the active cell the formula found in `G2` using Row Column notation

`R2C7` in Row Column notation essentially means that you want the cell in the **second row**, **seventh column**, which is `G7` in this case. 

`ActiveCell.FormulaR1C1 = "= 65 - R2C7`

#### Assign the active cell the formula `= 65 - the cell one to the left of the active cell`

`RC[-1]` in Row Column notation essentially means that you want the cell one column to the left of your active cell.

`ActiveCell.FormulaR1C1 = "= 65 - RC[-1]`

([Home](#table-of-contents))

## **Using Excel Functions**

### ***Math Functions***

([Home](#table-of-contents))
#### Find the AVERAGE value of the range from `A2` to `A26` and place it in cell `A28` 
`Range("A28").Value = Application.WorksheetFunction.Average(Range("A2:A26"))`

([Home](#table-of-contents))
#### Find the minimum of column `A`

`Application.WorksheetFunction.Min(Range("A2").EntireColumn.Select)`

You need to input a range inside the `Min()` brackets.

([Home](#table-of-contents))
#### Find the max of row `3`
`Application.WorksheetFunction.Min(Range("A3").EntireRow.Select)`

([Home](#table-of-contents))
#### Select all the values starting at `A2` down (`CTRL + DOWN`) in Excel, then sums the values 
`Range("A30").Value = Application.WorksheetFunction.Sum(Range("A2", Range("A2").End(xlDown)))`

#### Use `VLOOKUP` to look up the cell value in `E2` from the second column of reference table `A1:C20`

``` 
Dim CityVariable As String
CityVariable = Application.WorksheetFunction.VLookup(Range("E2"), Range("A1:C20"),2,False)
MsgBox("Employee " & Range("E2").Value & " is located in " & CityVariable)
```

![VLookup](/VBA/Images/vlookup_vba.png)

([Home](#table-of-contents))
## **User Interaction**

### ***Message Boxes***

#### A simple message box with the text "Learning is kewl! and "OK" Button

`MsgBox ("Learning is kewl!")`

This also works:

`MsgBox("Learning is kewl!", vkOKOnly)`

The `vbOKOnly` option specifies that only an OK button is available for users to click on.

([Home](#table-of-contents))

#### A message box with two lines of text on it

`MsgBox("This is line 1" & vbcrlf & "This is line 2")`

([Home](#table-of-contents))

#### A message box with the text "Learning is kewl!" and "Breaking News..." as the title

`MsgBox("Learning is kewl!", vkOKOnly, "Hello")`

([Home](#table-of-contents))

#### Read the text from cell `A1` and display it inside the message box

`MsgBox(Range("A1").Value, vbOKOnly, "Hello")`
([Home](#table-of-contents))

#### Ask the user if they would like to proceed or not in a message box

Keep in mind that when a user clicks `Yes` inside a message box, the actual response returned is `6`. When he/she clicks `No`, the string response returned is `7`. 

```
Dim UserResponse As String
UserResponse = MsgBox("Would you like to proceeed?", vbYesNo, "Proceed?")
If UserResponse = 6 Then
    MsgBox("The user's response was yes.")
Else
    MsgBox("The user's response was no.")
End If
```

([Home](#table-of-contents))
### ***Dialog Boxes***

#### Save one tab of a multi-tab workbook to a new file using the `Save As` dialog box

```
Sub SaveAFile()

Dim NewBook As Workbook
Dim FileName As String
Dim SheetName As String
Dim TheRange As Range

Set TheRange = Range("A1").CurrentRegion
SheetName = ActiveSheet.Name

Application.SheetsInNewWorkbook = 1
Set NewBook = Workbooks.Add
Application.SheetsInNewWorkbook = 3 

TheRange.Copy Destination:=Worksheets("Sheet1").Range("A1")
ActiveSheet.Name = SheetName
With Application.FileDialog(msoFileDialogSaveAs)
    .Title = "Please save the export"
    .Show
    If .SelectedItems.count = 0 Then
        MsgBox("A filename wasn't entered.")
        Newbook.Close savechanges:=False
        Exit Sub
    Else
    FileName = .SelectedItem(1)
    End If
End With

NewBook.SaveAs filename:=FileName
NewBook.Close
```
([Home](#table-of-contents))

### ***Input Boxes***

#### Ask a user for their first and last name in an input box and greet them with text in cell `C3`.

```
Sub AskNames()

Dim FirstName As String
Dim SecondName As String
Dim Greeting As String

FirstName = InputBox("What's our first name?", "Hello!")

SecondName = InputBox("What's your last name?", "Hello " & FirstName & "!")

Greeting = "I'm very happy to meet you, " & FirstName & " " & SecondName

Range("C3").Value = Greeting

End Sub
```
![Greeting](/VBA/Images/input_box_greeting.png)
([Home](#table-of-contents))

#### Ask user to enter their name, and validate that input was entered using `IF` condition

```
Sub InputBoxes()

    Dim FirstName As String

    FirstName = InputBox("Enter your name.", "Do it!")

    If FirstName = "" Then

    MsgBox("You didn't enter a name, dude!")

    Else

    MsgBox("Hello, " & FirstName)

End Sub
```

([Home](#table-of-contents))
#### Ask user to click a cell and return the address of that cell

```
Sub InputBox()

Dim ResponseFromUser As Range
Dim UserCellAddress As String

Set ResponseFromUser = Application.InputBox("Please select the cells you'd like to work on.", "Select Cells", Type:=8)

UserCellAddress = ResponseFromUser.Address

MsgBox("The range you clicked on is found at " & UserCellAddress)

End Sub
```
![Address](/VBA/Images/input_box_cell_address.png)

([Home](#table-of-contents))

#### Ask for a decimal number value in the Input Box and place a type check to validate

`InputBox("Enter a decimal:", Type:=1)`

([Home](#table-of-contents))
## **Loops**

### ***Do Loops***

#### Start at `A1` and move down until reaching an empty cell

```
Range("A1").Select

Do While ActiveCell.Value <> ""

    ActiveCell.Offset(1,0).Select

Loop
```
([Home](#table-of-contents))

#### Start at `A1` and move down until reaching a cell with a value less than `50`

```
Range("A1").Select

Do While ActiveCell.Value < 50

    ActiveCell.Offset(1,0).Select

Loop
```
([Home](#table-of-contents))

### ***For Next Loops***

#### Loop from `A1` down 20 cells

```
Dim LoopCounter As Integer

Sub ForLoopExample()
Range("A1").Select


For LoopCounter = 1 To 20

    ActiveCell.Offset(1,0).Select

Next

End Sub
```

([Home](#table-of-contents))

#### Loop from `A1` down `20` cells, with steps of `2`

```
Dim LoopCounter As Integer

Sub ForLoopStepExample()

Range("A1").Select

For LoopCounter = 1 To 20 Step 2

    ActiveCell.Offset(1,0).Select

Next

End Sub
```

([Home](#table-of-contents))

## **Logic and Conditions**

### ***If Statements***

#### Check if the active cell has a value of `5`
```
If ActiveCell.Value = 5 Then
    MsgBox("It's 5!)
End If
```

#### Check whether or not the user clicked `Yes` in a message box
```
YesOrNo = MsgBox("Click something...", vbYesNo)
If YesOrNo = vbYes Then
    MsgBox("You clicked yes!)
End If
```

### Select Case

#### Check the value of cell `A1` 

```
Select Case Range("A1").Value
    Case 0 To 10: Range("A2).Value = "A Small Number"
    Case 11 To 20: Range("A2).Value = "A Kinda Small Number"
    Case 21 To 30: Range("A2).Value = "A Medium Number"
    Case 31 To 40: Range("A2).Value = "A Large Number"
```
([Home](#table-of-contents))

# Debugging

Debugging refers to the process of removing errors in the code so that it runs and functions as desired. This can be a very frustrating process, but it's important to keep a few things in mind while debugging:

* **It's probably you**: that somehow doesn't run the code correctly. 99% of the time, it is **your fault** in the sense that something you wrote is not what it should be. I have only ever encountered maybe one instance ever where the source of the bug was a hardware issue.
* **Be scientific** about how you debug. First, try to replicate the error. If you can do that, you begin to pin down the source of the bug. Then, go **incrementally** alter variables and lines and observe the impact. Do not attempt to make a whole broad brush stroke of fixes at the same time, since you won't be able to identify what caused the change in behavior of your new code.