# 1. Table of Contents

<!-- TOC -->

- [1. Table of Contents](#1-table-of-contents)
- [2. **Formatting**](#2-formatting)
        - [2.0.1. ***Color***](#201-color)
            - [2.0.1.1. Set the background color of cell `A1` to red](#2011-set-the-background-color-of-cell-a1-to-red)
        - [2.0.2. **Font**](#202-font)
            - [2.0.2.1. Set the font in `A9` to bold](#2021-set-the-font-in-a9-to-bold)
            - [2.0.2.2. Set the font in `A2` to be regular](#2022-set-the-font-in-a2-to-be-regular)
            - [2.0.2.3. Set the cell `B4` to be both bold and italic:](#2023-set-the-cell-b4-to-be-both-bold-and-italic)
- [3. **Variables**](#3-variables)
    - [3.1. **Types of Variables**](#31-types-of-variables)
    - [3.2. ** Declaring Variables**](#32--declaring-variables)
            - [3.2.0.4. Force VBA to explicitly declare variables](#3204-force-vba-to-explicitly-declare-variables)
        - [3.2.1. ***Module Level***](#321-module-level)
            - [3.2.1.1. Declare a `Project` Level variable](#3211-declare-a-project-level-variable)
            - [3.2.1.2. Declare a `Module` Level variable](#3212-declare-a-module-level-variable)
            - [3.2.1.3. Declaring a String (text) variable, assign the variable the value of cell `A1`, and then assign the value of cell `C1` the value of the variable](#3213-declaring-a-string-text-variable-assign-the-variable-the-value-of-cell-a1-and-then-assign-the-value-of-cell-c1-the-value-of-the-variable)
- [4. **Basic Functions**](#4-basic-functions)
    - [**Working With New Sheets and Workbooks**](#working-with-new-sheets-and-workbooks)
        - [***Sheets***](#sheets)
            - [Save the name of the active sheet to a variable](#save-the-name-of-the-active-sheet-to-a-variable)
            - [Add a new sheet after the active sheet](#add-a-new-sheet-after-the-active-sheet)
            - [Change sheet orientation to landscape](#change-sheet-orientation-to-landscape)
    - [4.1. **Other**](#41-other)
        - [4.1.1. ***Formatting your code***](#411-formatting-your-code)
            - [4.1.1.1. Separating your code onto multiple lines](#4111-separating-your-code-onto-multiple-lines)
        - [4.1.2. ***Address of cells and ranges***](#412-address-of-cells-and-ranges)
            - [4.1.2.1. Find the address of the active cell](#4121-find-the-address-of-the-active-cell)
            - [4.1.2.2. Get the address of the last cell in the column `A`](#4122-get-the-address-of-the-last-cell-in-the-column-a)
        - [4.1.3. ***Clearing and Deleting Stuff***](#413-clearing-and-deleting-stuff)
            - [4.1.3.1. Clear (delete) the contents of column `G`](#4131-clear-delete-the-contents-of-column-g)
            - [4.1.3.2. Clear the contents of all cells](#4132-clear-the-contents-of-all-cells)
            - [4.1.3.3. Assign a range of cells to a variable](#4133-assign-a-range-of-cells-to-a-variable)
            - [4.1.3.4. Assigning Value to Cell](#4134-assigning-value-to-cell)
            - [4.1.3.5. Assign the value `"Yu Chen"` to a range of cells from `A1` to `D2`](#4135-assign-the-value-yu-chen-to-a-range-of-cells-from-a1-to-d2)
            - [4.1.3.6. Assign a value of `"Yu Chen"` to the variable `MyVariable`, and then assign this variable to cell `B2`](#4136-assign-a-value-of-yu-chen-to-the-variable-myvariable-and-then-assign-this-variable-to-cell-b2)
            - [4.1.3.7. Assigning Formula to Cell](#4137-assigning-formula-to-cell)
    - [4.2. **Selecting Things**](#42-selecting-things)
        - [4.2.1. ***Selecting Workbooks***](#421-selecting-workbooks)
            - [4.2.1.1. Activating the current workbook (where the code resides)](#4211-activating-the-current-workbook-where-the-code-resides)
            - [4.2.1.2. Activate a workbook to the name of `My Macro Book`](#4212-activate-a-workbook-to-the-name-of-my-macro-book)
            - [4.2.1.3. Activate the 2nd workbook (or specifically, the workbook in index position `2`)](#4213-activate-the-2nd-workbook-or-specifically-the-workbook-in-index-position-2)
        - [4.2.2. ***Selecting Cells***](#422-selecting-cells)
            - [4.2.2.1. **Select a single cell `A2`**](#4221-select-a-single-cell-a2)
            - [4.2.2.2. Select a group of cells that are not next to each other](#4222-select-a-group-of-cells-that-are-not-next-to-each-other)
            - [4.2.2.3. **Select the cells from `A2` to `B10`**](#4223-select-the-cells-from-a2-to-b10)
            - [4.2.2.4. Select the last row in column `A` of the dataset, and then moves `6` rows down](#4224-select-the-last-row-in-column-a-of-the-dataset-and-then-moves-6-rows-down)
            - [4.2.2.5. Select the last row in column `C` of the dataset, and then moves `6` rows down](#4225-select-the-last-row-in-column-c-of-the-dataset-and-then-moves-6-rows-down)
            - [4.2.2.6. Select the entire region of cells](#4226-select-the-entire-region-of-cells)
            - [4.2.2.7. Select entire column that `A2` is on (column `A`)](#4227-select-entire-column-that-a2-is-on-column-a)
            - [4.2.2.8. Select entire row that `A2` is on (row `2`)](#4228-select-entire-row-that-a2-is-on-row-2)
        - [4.2.3. ***Selecting Sheets***](#423-selecting-sheets)
            - [4.2.3.1. Select a sheet by tab name (`Sheets2`)](#4231-select-a-sheet-by-tab-name-sheets2)
            - [4.2.3.2. Select the next sheet in your workbook](#4232-select-the-next-sheet-in-your-workbook)
            - [4.2.3.3. Select the previous sheet in your workbook](#4233-select-the-previous-sheet-in-your-workbook)
        - [4.2.4. ***Selecting Worksheets***](#424-selecting-worksheets)
            - [4.2.4.1. **Select the `Task 2 Min and Max` tab**](#4241-select-the-task-2-min-and-max-tab)
        - [4.2.5. **2.1.4. Copying / Pasting Things**](#425-214-copying--pasting-things)
            - [4.2.5.1. Copy the value in cell `M3`](#4251-copy-the-value-in-cell-m3)
            - [4.2.5.2. Copy and paste value from cell `A1` to `B1` all in one line](#4252-copy-and-paste-value-from-cell-a1-to-b1-all-in-one-line)
            - [4.2.5.3. Assign the cell `M2` the value `10`. Assign the cell `M3` the formula `=SUM(M2,2)`, which should equal `12`. Copy this formula. Paste this formula from `M4` down to `M100`.](#4253-assign-the-cell-m2-the-value-10-assign-the-cell-m3-the-formula-summ22-which-should-equal-12-copy-this-formula-paste-this-formula-from-m4-down-to-m100)
    - [4.3. **Formulas**](#43-formulas)
            - [4.3.0.4. Assigns each cell from `D40` to `F40` the formula found in cell `F28`](#4304-assigns-each-cell-from-d40-to-f40-the-formula-found-in-cell-f28)
            - [Assigns each cell from `D40` to `F40` a particular formula](#assigns-each-cell-from-d40-to-f40-a-particular-formula)
            - [4.3.0.5. Assign each cell from `D40` to `F40` the formula found in cell `F28`](#4305-assign-each-cell-from-d40-to-f40-the-formula-found-in-cell-f28)
        - [4.3.1. ***Formulas with Row Column Notation***](#431-formulas-with-row-column-notation)
            - [4.3.1.1. Assign the active cell the formula found in `G2` using Row Column notation](#4311-assign-the-active-cell-the-formula-found-in-g2-using-row-column-notation)
            - [4.3.1.2. Assign the active cell the formula `= 65 - the cell one to the left of the active cell`](#4312-assign-the-active-cell-the-formula--65---the-cell-one-to-the-left-of-the-active-cell)
    - [4.4. **Using Excel Functions**](#44-using-excel-functions)
        - [4.4.1. ***Math Functions***](#441-math-functions)
            - [4.4.1.1. Find the AVERAGE value of the range from `A2` to `A26` and place it in cell `A28`](#4411-find-the-average-value-of-the-range-from-a2-to-a26-and-place-it-in-cell-a28)
            - [4.4.1.2. Find the minimum of column `A`](#4412-find-the-minimum-of-column-a)
            - [4.4.1.3. Find the max of row `3`](#4413-find-the-max-of-row-3)
            - [4.4.1.4. Select all the values starting at `A2` down (`CTRL + DOWN`) in Excel, then sums the values](#4414-select-all-the-values-starting-at-a2-down-ctrl--down-in-excel-then-sums-the-values)
            - [4.4.1.5. Use `VLOOKUP` to look up the cell value in `E2` from the second column of reference table `A1:C20`](#4415-use-vlookup-to-look-up-the-cell-value-in-e2-from-the-second-column-of-reference-table-a1c20)
    - [4.5. **User Interaction**](#45-user-interaction)
        - [4.5.1. ***Message Boxes***](#451-message-boxes)
            - [4.5.1.1. A simple message box with the text "Learning is kewl! and "OK" Button](#4511-a-simple-message-box-with-the-text-learning-is-kewl-and-ok-button)
            - [4.5.1.2. A message box with the text "Learning is kewl!" and "Breaking News..." as the title](#4512-a-message-box-with-the-text-learning-is-kewl-and-breaking-news-as-the-title)
            - [4.5.1.3. Read the text from cell `A1` and display it inside the message box](#4513-read-the-text-from-cell-a1-and-display-it-inside-the-message-box)
            - [4.5.1.4. Ask the user if they would like to proceed or not in a message box](#4514-ask-the-user-if-they-would-like-to-proceed-or-not-in-a-message-box)
        - [4.5.2. ***Input Boxes***](#452-input-boxes)
            - [4.5.2.1. Ask a user for their first and last name in an input box and greet them with text in cell `C3`.](#4521-ask-a-user-for-their-first-and-last-name-in-an-input-box-and-greet-them-with-text-in-cell-c3)
            - [4.5.2.2. Ask user to click a cell and return the address of that cell](#4522-ask-user-to-click-a-cell-and-return-the-address-of-that-cell)
- [5. **Loops**](#5-loops)
        - [5.0.3. ***Do Loops***](#503-do-loops)
            - [5.0.3.1. Start at `A1` and move down until reaching an empty cell](#5031-start-at-a1-and-move-down-until-reaching-an-empty-cell)
            - [5.0.3.2. Start at `A1` and move down until reaching a cell with a value less than `50`](#5032-start-at-a1-and-move-down-until-reaching-a-cell-with-a-value-less-than-50)
        - [5.0.4. ***For Next Loops***](#504-for-next-loops)
            - [5.0.4.1. Loop from `A1` down 20 cells](#5041-loop-from-a1-down-20-cells)
            - [5.0.4.2. Loop from `A1` down `20` cells, with steps of `2`](#5042-loop-from-a1-down-20-cells-with-steps-of-2)

<!-- /TOC -->

# 2. **Formatting**

### 2.0.1. ***Color***

#### 2.0.1.1. Set the background color of cell `A1` to red

`Range("A1").Interior.Color = vbRed`

### 2.0.2. **Font**

#### 2.0.2.1. Set the font in `A9` to bold
`Range("A9").Font.Bold = True`

You can also set it this way:

`Range("A9").Font.FontStyle = "Bold"`
([Home](#1-table-of-contents))
#### 2.0.2.2. Set the font in `A2` to be regular
`Range("A9").Font.Bold = False`

You can also set it this way:

`Range("A9").Font.FontStyle = "Regular"`
([Home](#1-table-of-contents))
#### 2.0.2.3. Set the cell `B4` to be both bold and italic:

`Range("B4").Font.FontStyle = "Bold italic"`
([Home](#1-table-of-contents))
# 3. **Variables**

Here are some general naming conventions when declaring and using variables:

- must be less than **255** characters
- must only use letters, numbers, or underscores (**no spaces**!)
- cannot begin with a number

There are different levels of **Scope** for a variable. Scope means what context the variable can be used.

- **Procedure Level**: this variable can only be used inside the subroutine that it has been declared in. They are declared as using `Dim`.
- **Module Level**: this variable can be used in all the subroutines in the current module. You declare the variable the same way as the **Procedure Level**, but place it before all of the `Sub()` declarations. 
- **Project Level**: this can be used anywhere in the current project. You declare this type of variable with the keyword `Public`.

## 3.1. **Types of Variables**

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
|  **Variant**  | Any of the above. | At least 16 bytes.

([Home](#1-table-of-contents))
## 3.2. ** Declaring Variables**

#### 3.2.0.4. Force VBA to explicitly declare variables

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

([Home](#1-table-of-contents))
### 3.2.1. ***Module Level***

#### 3.2.1.1. Declare a `Project` Level variable

`Public myName As String`

Project variables are declared using the keyword `Public`. It is available to all subroutines in the entire project.

#### 3.2.1.2. Declare a `Module` Level variable

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

#### 3.2.1.3. Declaring a String (text) variable, assign the variable the value of cell `A1`, and then assign the value of cell `C1` the value of the variable

```
Dim myNewStringVar As String
Range("A1").Select
myNewStringVar = ActiveCell.Value
Range("C1").Value = myNewStringVar
```

([Home](#1-table-of-contents))
# 4. **Basic Functions**

## **Working With New Sheets and Workbooks**

### ***Sheets***

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

## 4.1. **Other**

### 4.1.1. ***Formatting your code***

#### 4.1.1.1. Separating your code onto multiple lines

Sometimes a particular function in VBA will take up a lot of space to write. One way to keep your code neat is to break it up into multiple lines. In order to tell VBA that you'd like to break up  your code into multiple lines, using ` _ ` to separate the code.

Here is an example:
` Temp = Application.WorksheetFunction.VLookup(Range("C2"), Ranage("A4:B200"),2,False)`

can be broken up into

```
Temp = Application.WorksheetFunction.VLookup _ 
    (Range("C2"), Range("A4:B200), 2, False)

```

([Home](#1-table-of-contents))

### 4.1.2. ***Address of cells and ranges***

#### 4.1.2.1. Find the address of the active cell
`ActiveCell.Address`

([Home](#1-table-of-contents))
#### 4.1.2.2. Get the address of the last cell in the column `A`
`Range("A1").End(xlDown).Address`

([Home](#1-table-of-contents))
### 4.1.3. ***Clearing and Deleting Stuff***

#### 4.1.3.1. Clear (delete) the contents of column `G`
`Range("G:G").ClearContents`

([Home](#1-table-of-contents))
#### 4.1.3.2. Clear the contents of all cells
`Cells.ClearContents`

([Home](#1-table-of-contents))
#### 4.1.3.3. Assign a range of cells to a variable
`MyRange = Range("A1:B20")`

([Home](#1-table-of-contents))
#### 4.1.3.4. Assigning Value to Cell
`Range("M2").Value = 10`

Note here that when you are entering in a number, you do not need to put in quotation marks! If you were to write this instead:

`Range("M2").Value = "10"`

You would get a string value, not a integer (number) value.

([Home](#1-table-of-contents))
#### 4.1.3.5. Assign the value `"Yu Chen"` to a range of cells from `A1` to `D2`
`Range("A1:D2").Value = "Yu Chen"`

([Home](#1-table-of-contents))

#### 4.1.3.6. Assign a value of `"Yu Chen"` to the variable `MyVariable`, and then assign this variable to cell `B2`
```
MyVariable = "Yu Chen"
Range("B2").Value = MyVariable
```
([Home](#1-table-of-contents))

#### 4.1.3.7. Assigning Formula to Cell
This assigns the Excel formula to `M3` (take the value of `M2` and add `2` to it.)

`Range("M3").Formula = =SUM(M2,2)`

([Home](#1-table-of-contents))
## 4.2. **Selecting Things**

([Home](#1-table-of-contents))
### 4.2.1. ***Selecting Workbooks***

([Home](#1-table-of-contents))
#### 4.2.1.1. Activating the current workbook (where the code resides)
`ThisWorkbook.Activate`

([Home](#1-table-of-contents))
#### 4.2.1.2. Activate a workbook to the name of `My Macro Book`
`Workbooks("My Work Book").Activate`

([Home](#1-table-of-contents))
#### 4.2.1.3. Activate the 2nd workbook (or specifically, the workbook in index position `2`)
`Workbooks(2).Activate`

([Home](#1-table-of-contents))
### 4.2.2. ***Selecting Cells***

([Home](#1-table-of-contents))
#### 4.2.2.1. **Select a single cell `A2`**

`Range("A2").Select`

([Home](#1-table-of-contents))
#### 4.2.2.2. Select a group of cells that are not next to each other
`Range("B2,C8,E9").Select`

`Range("B2,C8,E9:E20").Select`

([Home](#1-table-of-contents))
#### 4.2.2.3. **Select the cells from `A2` to `B10`**

`Range("A2", "B10").Select`

([Home](#1-table-of-contents))
#### 4.2.2.4. Select the last row in column `A` of the dataset, and then moves `6` rows down 
`Range("A1").End(xlDown).Offset(6, 0).Select`

([Home](#1-table-of-contents))
#### 4.2.2.5. Select the last row in column `C` of the dataset, and then moves `6` rows down
`Range("C1").End(xlDown).Offset(-6, 0).Select`

([Home](#1-table-of-contents))
#### 4.2.2.6. Select the entire region of cells
This is equivalent to hitting `CTRL + SHIFT + DOWN + RIGHT` on your keyboard:

`ActiveCell.CurrentRegion.Select`

This is also equivalent to the following command:

`Range("A2", Range("A2").End(xlDown).End(xlToRight)).Select`

([Home](#1-table-of-contents))
#### 4.2.2.7. Select entire column that `A2` is on (column `A`)

`Range("A2").EntireColumn.Select`

([Home](#1-table-of-contents))
#### 4.2.2.8. Select entire row that `A2` is on (row `2`) 
`Range("A2").EntireRow.Select`

([Home](#1-table-of-contents))
### 4.2.3. ***Selecting Sheets***

A `Sheet` and a `Worksheet` are related, but cannot be used interchangeably. A `Sheet` is any Excel sheet, whereas a `Worksheet` is only a regular Excel worksheet. For example, a chart is a `Sheet` but is not a `Worksheet`.

([Home](#1-table-of-contents))
#### 4.2.3.1. Select a sheet by tab name (`Sheets2`)
`Sheets("Sheets2").Select`

([Home](#1-table-of-contents))
#### 4.2.3.2. Select the next sheet in your workbook

`ActiveSheet.Next.Select`

([Home](#1-table-of-contents))
#### 4.2.3.3. Select the previous sheet in your workbook

`ActiveSheet.Previous.Select`

([Home](#1-table-of-contents))
### 4.2.4. ***Selecting Worksheets***

#### 4.2.4.1. **Select the `Task 2 Min and Max` tab**
`Worksheets("Task 2 Min and Max").Activate`

([Home](#1-table-of-contents))
### 4.2.5. **2.1.4. Copying / Pasting Things**

([Home](#1-table-of-contents))
#### 4.2.5.1. Copy the value in cell `M3`
`Range("M3").Copy`

([Home](#1-table-of-contents))
#### 4.2.5.2. Copy and paste value from cell `A1` to `B1` all in one line
`Range("A1").Copy Range("B1")`

Note that this pastes all the formatting as well, so if you had a bolded cell in `A1`, you'll also have a bolded cell in `B1`.

([Home](#1-table-of-contents))
#### 4.2.5.3. Assign the cell `M2` the value `10`. Assign the cell `M3` the formula `=SUM(M2,2)`, which should equal `12`. Copy this formula. Paste this formula from `M4` down to `M100`.

```
Range("M2").Value = 10 
Range("M3").Formula = "=SUM(M2,2)"
Range("M3").Copy 
Range("M4:D100").PasteSpecial
```
([Home](#1-table-of-contents))

## 4.3. **Formulas**

#### 4.3.0.4. Assigns each cell from `D40` to `F40` the formula found in cell `F28` 
`Range("D40:F40").Formula = Range("F28").Formula`

([Home](#1-table-of-contents))
#### Assigns each cell from `D40` to `F40` a particular formula

`Range("D40:F40").Formula = "=65 - A2"`

Please note that when you do this, the formula is `F28` is interpreted **relatively**. For example, if your formula is `= 65 - A2`, the first row will have `= 65 - A2`, but the sec ond row will have `= 65 - A3` and the third row will have `= 65 - A4`.

([Home](#1-table-of-contents))
#### 4.3.0.5. Assign each cell from `D40` to `F40` the formula found in cell `F28`
`Range("D40:F40").Formula = Range("F28").Formula`

### 4.3.1. ***Formulas with Row Column Notation***

#### 4.3.1.1. Assign the active cell the formula found in `G2` using Row Column notation

`R2C7` in Row Column notation essentially means that you want the cell in the **second row**, **seventh column**, which is `G7` in this case. 

`ActiveCell.FormulaR1C1 = "= 65 - R2C7`

#### 4.3.1.2. Assign the active cell the formula `= 65 - the cell one to the left of the active cell`

`RC[-1]` in Row Column notation essentially means that you want the cell one column to the left of your active cell.

`ActiveCell.FormulaR1C1 = "= 65 - RC[-1]`

([Home](#1-table-of-contents))

## 4.4. **Using Excel Functions**

### 4.4.1. ***Math Functions***

([Home](#1-table-of-contents))
#### 4.4.1.1. Find the AVERAGE value of the range from `A2` to `A26` and place it in cell `A28` 
`Range("A28").Value = Application.WorksheetFunction.Average(Range("A2:A26"))`

([Home](#1-table-of-contents))
#### 4.4.1.2. Find the minimum of column `A`

`Application.WorksheetFunction.Min(Range("A2").EntireColumn.Select)`

You need to input a range inside the `Min()` brackets.

([Home](#1-table-of-contents))
#### 4.4.1.3. Find the max of row `3`
`Application.WorksheetFunction.Min(Range("A3").EntireRow.Select)`

([Home](#1-table-of-contents))
#### 4.4.1.4. Select all the values starting at `A2` down (`CTRL + DOWN`) in Excel, then sums the values 
`Range("A30").Value = Application.WorksheetFunction.Sum(Range("A2", Range("A2").End(xlDown)))`

#### 4.4.1.5. Use `VLOOKUP` to look up the cell value in `E2` from the second column of reference table `A1:C20`

``` 
Dim CityVariable As String
CityVariable = Application.WorksheetFunction.VLookup(Range("E2"), Range("A1:C20"),2,False)
MsgBox("Employee " & Range("E2").Value & " is located in " & CityVariable)
```

![VLookup](/VBA/Images/vlookup_vba.png)

([Home](#1-table-of-contents))
## 4.5. **User Interaction**

### 4.5.1. ***Message Boxes***

#### 4.5.1.1. A simple message box with the text "Learning is kewl! and "OK" Button

`MsgBox ("Learning is kewl!")`

This also works:

`MsgBox("Learning is kewl!", vkOKOnly)`

([Home](#1-table-of-contents))

#### 4.5.1.2. A message box with the text "Learning is kewl!" and "Breaking News..." as the title

`MsgBox("Learning is kewl!", vkOKOnly, "Hello")`

([Home](#1-table-of-contents))

#### 4.5.1.3. Read the text from cell `A1` and display it inside the message box

`MsgBox(Range("A1").Value, vbOKOnly, "Hello")`

#### 4.5.1.4. Ask the user if they would like to proceed or not in a message box

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

([Home](#1-table-of-contents))

### 4.5.2. ***Input Boxes***

#### 4.5.2.1. Ask a user for their first and last name in an input box and greet them with text in cell `C3`.

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
([Home](#1-table-of-contents))

#### 4.5.2.2. Ask user to click a cell and return the address of that cell

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

([Home](#1-table-of-contents))

# 5. **Loops**

### 5.0.3. ***Do Loops***

#### 5.0.3.1. Start at `A1` and move down until reaching an empty cell

```
Range("A1").Select

Do While ActiveCell.Value <> ""

    ActiveCell.Offset(1,0).Select

Loop
```
([Home](#1-table-of-contents))

#### 5.0.3.2. Start at `A1` and move down until reaching a cell with a value less than `50`

```
Range("A1").Select

Do While ActiveCell.Value < 50

    ActiveCell.Offset(1,0).Select

Loop
```
([Home](#1-table-of-contents))

### 5.0.4. ***For Next Loops***

#### 5.0.4.1. Loop from `A1` down 20 cells

```
Dim LoopCounter As Integer

Sub ForLoopExample()
Range("A1").Select


For LoopCounter = 1 To 20

    ActiveCell.Offset(1,0).Select

Next

End Sub
```

([Home](#1-table-of-contents))

#### 5.0.4.2. Loop from `A1` down `20` cells, with steps of `2`

```
Dim LoopCounter As Integer

Sub ForLoopStepExample()

Range("A1").Select

For LoopCounter = 1 To 20 Step 2

    ActiveCell.Offset(1,0).Select

Next

End Sub
```

([Home](#1-table-of-contents))
