# 1. Table of Contents

<!-- TOC -->

- [1. Table of Contents](#1-table-of-contents)
- [2. **Formatting**](#2-formatting)
        - [2.0.1. **Color**](#201-color)
        - [2.0.2. **Font**](#202-font)
            - [2.0.2.1. Set the font in `A9` to bold](#2021-set-the-font-in-a9-to-bold)
            - [2.0.2.2. Set the font in `A2` to be regular](#2022-set-the-font-in-a2-to-be-regular)
            - [2.0.2.3. Set the cell `B4` to be both bold and italic:](#2023-set-the-cell-b4-to-be-both-bold-and-italic)
- [3. Variables](#3-variables)
    - [3.1. **Types of Variables**](#31-types-of-variables)
    - [3.2. **3.2. Declaring Variables**](#32-32-declaring-variables)
            - [3.2.0.4. Force VBA to explicitly declare variables](#3204-force-vba-to-explicitly-declare-variables)
        - [3.2.1. **Module Level**](#321-module-level)
            - [3.2.1.1. Declare a `Project` Level variable](#3211-declare-a-project-level-variable)
            - [3.2.1.2. Declare a `Module` Level variable](#3212-declare-a-module-level-variable)
            - [3.2.1.3. Declaring a String (text) variable, assign the variable the value of cell `A1`, and then assign the value of cell `C1` the value of the variable](#3213-declaring-a-string-text-variable-assign-the-variable-the-value-of-cell-a1-and-then-assign-the-value-of-cell-c1-the-value-of-the-variable)
- [4. **Basic Functions**](#4-basic-functions)
        - [4.0.2. **Other**](#402-other)
            - [4.0.2.1. Clear (delete) the contents of column G](#4021-clear-delete-the-contents-of-column-g)
            - [4.0.2.2. Assigning Value to Cell](#4022-assigning-value-to-cell)
            - [4.0.2.3. Assign the value "Yu Chen" to a range of cells from `A1` to `D2`](#4023-assign-the-value-yu-chen-to-a-range-of-cells-from-a1-to-d2)
            - [4.0.2.4. Assign a value of `"Yu Chen"` to the variable `MyVariable`, and then assign this variable to cell `B2`](#4024-assign-a-value-of-yu-chen-to-the-variable-myvariable-and-then-assign-this-variable-to-cell-b2)
            - [4.0.2.5. Assigning Formula to Cell](#4025-assigning-formula-to-cell)
    - [4.1. **Selecting Things**](#41-selecting-things)
        - [4.1.1. **Selecting Workbooks**](#411-selecting-workbooks)
            - [4.1.1.1. **Activating the current workbook (where the code resides)**](#4111-activating-the-current-workbook-where-the-code-resides)
            - [4.1.1.2. **Activate a workbook to the name of `My Macro Book`**](#4112-activate-a-workbook-to-the-name-of-my-macro-book)
            - [4.1.1.3. **Activate the 2nd workbook (or specifically, the workbook in index position `2`)**](#4113-activate-the-2nd-workbook-or-specifically-the-workbook-in-index-position-2)
        - [4.1.2. **Selecting Cells**](#412-selecting-cells)
            - [4.1.2.1. **Select a single cell `A2`**](#4121-select-a-single-cell-a2)
            - [4.1.2.2. **Select a group of cells that are not next to each other **](#4122-select-a-group-of-cells-that-are-not-next-to-each-other-)
            - [4.1.2.3. **Select the cells from `A2` to `B10`**](#4123-select-the-cells-from-a2-to-b10)
            - [4.1.2.4. **Select the last row in column `A` of the dataset, and then moves `6` rows down**](#4124-select-the-last-row-in-column-a-of-the-dataset-and-then-moves-6-rows-down)
            - [4.1.2.5. **Select the last row in column `C` of the dataset, and then moves `6` rows down**](#4125-select-the-last-row-in-column-c-of-the-dataset-and-then-moves-6-rows-down)
            - [4.1.2.6. **Select the entire region of cells**](#4126-select-the-entire-region-of-cells)
            - [4.1.2.7. **Select entire column that `A2` is on (column `A`)**](#4127-select-entire-column-that-a2-is-on-column-a)
            - [4.1.2.8. Select entire row that `A2` is on (row `2`)](#4128-select-entire-row-that-a2-is-on-row-2)
        - [4.1.3. **Selecting Sheets**](#413-selecting-sheets)
            - [4.1.3.1. Select a sheet by tab name (`Sheets2`)](#4131-select-a-sheet-by-tab-name-sheets2)
            - [4.1.3.2. Select the next sheet in your workbook](#4132-select-the-next-sheet-in-your-workbook)
            - [4.1.3.3. Select the previous sheet in your workbook](#4133-select-the-previous-sheet-in-your-workbook)
        - [4.1.4. **Selecting Worksheets**](#414-selecting-worksheets)
            - [4.1.4.1. **Select the `Task 2 Min and Max` tab**](#4141-select-the-task-2-min-and-max-tab)
        - [4.1.5. **2.1.4. Copying / Pasting Things**](#415-214-copying--pasting-things)
            - [4.1.5.1. Copy the value in cell `M3`](#4151-copy-the-value-in-cell-m3)
            - [4.1.5.2. Copy and paste value from cell `A1` to `B1` all in one line](#4152-copy-and-paste-value-from-cell-a1-to-b1-all-in-one-line)
            - [4.1.5.3. Assign the cell `M2` the value `10`. Assign the cell `M3` the formula `=SUM(M2,2)`, which should equal `12`. Copy this formula. Paste this formula from `M4` down to `M100`.](#4153-assign-the-cell-m2-the-value-10-assign-the-cell-m3-the-formula-summ22-which-should-equal-12-copy-this-formula-paste-this-formula-from-m4-down-to-m100)
    - [4.2. **Formulas**](#42-formulas)
            - [4.2.0.4. Assigns each cell from `D40` to `F40` the formula found in cell `F28`](#4204-assigns-each-cell-from-d40-to-f40-the-formula-found-in-cell-f28)
            - [4.2.0.5. Assign each cell from `D40` to `F40` the formula found in cell `F28`](#4205-assign-each-cell-from-d40-to-f40-the-formula-found-in-cell-f28)
    - [4.3. **2.3. Functions**](#43-23-functions)
        - [4.3.1. **Math Functions**](#431-math-functions)
            - [4.3.1.1. Find the AVERAGE value of the range from `A2` to `A26` and place it in cell `A28`](#4311-find-the-average-value-of-the-range-from-a2-to-a26-and-place-it-in-cell-a28)
            - [4.3.1.2. Find the minimum of column `A`](#4312-find-the-minimum-of-column-a)
            - [4.3.1.3. Find the max of row `3`](#4313-find-the-max-of-row-3)
            - [4.3.1.4. Select all the values starting at `A2` down (`CTRL + DOWN`) in Excel, then sums the values](#4314-select-all-the-values-starting-at-a2-down-ctrl--down-in-excel-then-sums-the-values)
    - [4.4. **User Interaction**](#44-user-interaction)
        - [4.4.1. **Message Boxes**](#441-message-boxes)
            - [4.4.1.1. A simple message box with the text "Learning is kewl! and "OK" Button](#4411-a-simple-message-box-with-the-text-learning-is-kewl-and-ok-button)
            - [Ask a user for their first and last name in a message box and greet them with text in cell `C3`.](#ask-a-user-for-their-first-and-last-name-in-a-message-box-and-greet-them-with-text-in-cell-c3)
- [5. **Loops**](#5-loops)
        - [5.0.2. **Do Loops**](#502-do-loops)
            - [5.0.2.1. Start at `A1` and move down until reaching an empty cell](#5021-start-at-a1-and-move-down-until-reaching-an-empty-cell)
            - [5.0.2.2. Start at `A1` and move down until reaching a cell with a value less than `50`](#5022-start-at-a1-and-move-down-until-reaching-a-cell-with-a-value-less-than-50)
        - [5.0.3. **For Next Loops**](#503-for-next-loops)
            - [5.0.3.1. Loop from `A1` down 20 cells](#5031-loop-from-a1-down-20-cells)
            - [5.0.3.2. Loop from `A1` down `20` cells, with steps of `2`](#5032-loop-from-a1-down-20-cells-with-steps-of-2)

<!-- /TOC -->

# 2. **Formatting**

### 2.0.1. **Color**


### 2.0.2. **Font**

#### 2.0.2.1. Set the font in `A9` to bold
`Range("A9").Font.Bold = True`

You can also set it this way:

`Range("A9").Font.FontStyle = "Bold"`

#### 2.0.2.2. Set the font in `A2` to be regular
`Range("A9").Font.Bold = False`

You can also set it this way:

`Range("A9").Font.FontStyle = "Regular"`

#### 2.0.2.3. Set the cell `B4` to be both bold and italic:

`Range("B4").Font.FontStyle = "Bold italic"`

# 3. Variables

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

## 3.2. **3.2. Declaring Variables**

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
### 3.2.1. **Module Level**

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

### 4.0.2. **Other**

#### 4.0.2.1. Clear (delete) the contents of column G
`Range("G:G").ClearContents`


([Home](#1-table-of-contents))

#### 4.0.2.2. Assigning Value to Cell
`Range("M2").Value = 10`

Note here that when you are entering in a number, you do not need to put in quotation marks! If you were to write this instead:

`Range("M2").Value = "10"`

You would get a string value, not a integer (number) value.

([Home](#1-table-of-contents))

#### 4.0.2.3. Assign the value "Yu Chen" to a range of cells from `A1` to `D2`
`Range("A1:D2").Value = "Yu Chen"`

([Home](#1-table-of-contents))

#### 4.0.2.4. Assign a value of `"Yu Chen"` to the variable `MyVariable`, and then assign this variable to cell `B2`
```
MyVariable = "Yu Chen"
Range("B2").Value = MyVariable
```
([Home](#1-table-of-contents))

#### 4.0.2.5. Assigning Formula to Cell
This assigns the Excel formula to `M3` (take the value of `M2` and add `2` to it.)

`Range("M3").Formula = =SUM(M2,2)`

([Home](#1-table-of-contents))
## 4.1. **Selecting Things**

([Home](#1-table-of-contents))
### 4.1.1. **Selecting Workbooks**

([Home](#1-table-of-contents))
#### 4.1.1.1. **Activating the current workbook (where the code resides)**
`ThisWorkbook.Activate`

([Home](#1-table-of-contents))
#### 4.1.1.2. **Activate a workbook to the name of `My Macro Book`**
`Workbooks("My Work Book").Activate`

([Home](#1-table-of-contents))
#### 4.1.1.3. **Activate the 2nd workbook (or specifically, the workbook in index position `2`)**
`Workbooks(2).Activate`

([Home](#1-table-of-contents))
### 4.1.2. **Selecting Cells**

([Home](#1-table-of-contents))
#### 4.1.2.1. **Select a single cell `A2`**

`Range("A2").Select`

([Home](#1-table-of-contents))
#### 4.1.2.2. **Select a group of cells that are not next to each other **
`Range("B2,C8,E9").Select`

`Range("B2,C8,E9:E20").Select`

([Home](#1-table-of-contents))
#### 4.1.2.3. **Select the cells from `A2` to `B10`**

`Range("A2", "B10").Select`

([Home](#1-table-of-contents))
#### 4.1.2.4. **Select the last row in column `A` of the dataset, and then moves `6` rows down** 
`Range("A1").End(xlDown).Offset(6, 0).Select`

([Home](#1-table-of-contents))
#### 4.1.2.5. **Select the last row in column `C` of the dataset, and then moves `6` rows down** 
`Range("C1").End(xlDown).Offset(-6, 0).Select`

([Home](#1-table-of-contents))
#### 4.1.2.6. **Select the entire region of cells** 
This is equivalent to hitting `CTRL + SHIFT + DOWN + RIGHT` on your keyboard:

`ActiveCell.CurrentRegion.Select`

This is also equivalent to the following command:

`Range("A2", Range("A2").End(xlDown).End(xlToRight)).Select`

([Home](#1-table-of-contents))
#### 4.1.2.7. **Select entire column that `A2` is on (column `A`)**

`Range("A2").EntireColumn.Select`

([Home](#1-table-of-contents))
#### 4.1.2.8. Select entire row that `A2` is on (row `2`) 
`Range("A2").EntireRow.Select`

([Home](#1-table-of-contents))
### 4.1.3. **Selecting Sheets**

A `Sheet` and a `Worksheet` are related, but cannot be used interchangeably. A `Sheet` is any Excel sheet, whereas a `Worksheet` is only a regular Excel worksheet. For example, a chart is a `Sheet` but is not a `Worksheet`.

([Home](#1-table-of-contents))
#### 4.1.3.1. Select a sheet by tab name (`Sheets2`)
`Sheets("Sheets2").Select`

([Home](#1-table-of-contents))
#### 4.1.3.2. Select the next sheet in your workbook

`ActiveSheet.Next.Select`

([Home](#1-table-of-contents))
#### 4.1.3.3. Select the previous sheet in your workbook

`ActiveSheet.Previous.Select`

([Home](#1-table-of-contents))
### 4.1.4. **Selecting Worksheets**

#### 4.1.4.1. **Select the `Task 2 Min and Max` tab**
`Worksheets("Task 2 Min and Max").Activate`

([Home](#1-table-of-contents))
### 4.1.5. **2.1.4. Copying / Pasting Things**

([Home](#1-table-of-contents))
#### 4.1.5.1. Copy the value in cell `M3`
`Range("M3").Copy`

([Home](#1-table-of-contents))
#### 4.1.5.2. Copy and paste value from cell `A1` to `B1` all in one line
`Range("A1").Copy Range("B1")`

Note that this pastes all the formatting as well, so if you had a bolded cell in `A1`, you'll also have a bolded cell in `B1`.

([Home](#1-table-of-contents))
#### 4.1.5.3. Assign the cell `M2` the value `10`. Assign the cell `M3` the formula `=SUM(M2,2)`, which should equal `12`. Copy this formula. Paste this formula from `M4` down to `M100`.

```
Range("M2").Value = 10 
Range("M3").Formula = "=SUM(M2,2)"
Range("M3").Copy 
Range("M4:D100").PasteSpecial
```
([Home](#1-table-of-contents))

## 4.2. **Formulas**

#### 4.2.0.4. Assigns each cell from `D40` to `F40` the formula found in cell `F28` 
`Range("D40:F40").Formula = Range("F28").Formula`

([Home](#1-table-of-contents))
#### 4.2.0.5. Assign each cell from `D40` to `F40` the formula found in cell `F28`
`Range("D40:F40").Formula = Range("F28").Formula`

([Home](#1-table-of-contents))

## 4.3. **2.3. Functions**

### 4.3.1. **Math Functions**

([Home](#1-table-of-contents))
#### 4.3.1.1. Find the AVERAGE value of the range from `A2` to `A26` and place it in cell `A28` 
`Range("A28").Value = Application.WorksheetFunction.Average(Range("A2:A26"))`

([Home](#1-table-of-contents))
#### 4.3.1.2. Find the minimum of column `A`

`Application.WorksheetFunction.Min(Range("A2").EntireColumn.Select)`

You need to input a range inside the `Min()` brackets.

([Home](#1-table-of-contents))
#### 4.3.1.3. Find the max of row `3`
`Application.WorksheetFunction.Min(Range("A3").EntireRow.Select)`

([Home](#1-table-of-contents))
#### 4.3.1.4. Select all the values starting at `A2` down (`CTRL + DOWN`) in Excel, then sums the values 
`Range("A30").Value = Application.WorksheetFunction.Sum(Range("A2", Range("A2").End(xlDown)))`

([Home](#1-table-of-contents))
## 4.4. **User Interaction**

### 4.4.1. **Message Boxes**

#### 4.4.1.1. A simple message box with the text "Learning is kewl! and "OK" Button

`MsgBox ("Learning is kewl!")`

([Home](#1-table-of-contents))

#### Ask a user for their first and last name in a message box and greet them with text in cell `C3`.

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

# 5. **Loops**

### 5.0.2. **Do Loops**

#### 5.0.2.1. Start at `A1` and move down until reaching an empty cell

```
Range("A1").Select

Do While ActiveCell.Value <> ""

    ActiveCell.Offset(1,0).Select

Loop
```
([Home](#1-table-of-contents))

#### 5.0.2.2. Start at `A1` and move down until reaching a cell with a value less than `50`

```
Range("A1").Select

Do While ActiveCell.Value < 50

    ActiveCell.Offset(1,0).Select

Loop
```
([Home](#1-table-of-contents))

### 5.0.3. **For Next Loops**

#### 5.0.3.1. Loop from `A1` down 20 cells

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

#### 5.0.3.2. Loop from `A1` down `20` cells, with steps of `2`

```
Dim LoopCounter As Integer

Sub ForLoopStepExample()

Range("A1").Select

For LoopCounter = 1 To 20 Step 2

    ActiveCell.Offset(1,0).Select

Next

End Sub
```