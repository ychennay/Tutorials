# 1. Table of Contents

<!-- TOC -->

- [1. Table of Contents](#1-table-of-contents)
- [2. **2. Basic Functions**](#2-2-basic-functions)
        - [2.0.1. Other](#201-other)
            - [2.0.1.1. Assigning Value to Cell ([Home](#1-table-of-contents))](#2011-assigning-value-to-cell-home1-table-of-contents)
            - [2.0.1.2. Assign the value "Yu Chen" to a range of cells from `A1` to `D2` ([Home](#1-table-of-contents))](#2012-assign-the-value-yu-chen-to-a-range-of-cells-from-a1-to-d2-home1-table-of-contents)
            - [2.0.1.3. Assign a value of `"Yu Chen"` to the variable `MyVariable`, and then assign this variable to cell `B2` ([Home](#1-table-of-contents))](#2013-assign-a-value-of-yu-chen-to-the-variable-myvariable-and-then-assign-this-variable-to-cell-b2-home1-table-of-contents)
            - [2.0.1.4. Assigning Formula to Cell](#2014-assigning-formula-to-cell)
    - [2.1. **2.1. Selecting Things**](#21-21-selecting-things)
        - [2.1.1. **Selecting Workbooks**](#211-selecting-workbooks)
            - [2.1.1.1. Activating the current workbook (where the code resides)](#2111-activating-the-current-workbook-where-the-code-resides)
            - [2.1.1.2. Activate a workbook to the name of `My Macro Book`](#2112-activate-a-workbook-to-the-name-of-my-macro-book)
            - [2.1.1.3. Activate the 2nd workbook (or specifically, the workbook in index position `2`)](#2113-activate-the-2nd-workbook-or-specifically-the-workbook-in-index-position-2)
        - [2.1.2. **2.1.1. Selecting Cells**](#212-211-selecting-cells)
            - [2.1.2.1. Select a single cell `A2`](#2121-select-a-single-cell-a2)
            - [2.1.2.2. Select the cells from `A2` to `B10`](#2122-select-the-cells-from-a2-to-b10)
            - [2.1.2.3. Select the last row in column `A` of the dataset, and then moves `6` rows down](#2123-select-the-last-row-in-column-a-of-the-dataset-and-then-moves-6-rows-down)
            - [2.1.2.4. Select the last row in column `C` of the dataset, and then moves `6` rows down](#2124-select-the-last-row-in-column-c-of-the-dataset-and-then-moves-6-rows-down)
            - [2.1.2.5. Select the **entire region of cells**](#2125-select-the-entire-region-of-cells)
            - [2.1.2.6. Select entire column that `A2` is on (column `A`)](#2126-select-entire-column-that-a2-is-on-column-a)
            - [2.1.2.7. Select entire row that `A2` is on (row `2`)](#2127-select-entire-row-that-a2-is-on-row-2)
        - [2.1.3. **2.1.2. Selecting Sheets**](#213-212-selecting-sheets)
            - [2.1.3.1. Select a sheet by tab name (`Sheets2`)](#2131-select-a-sheet-by-tab-name-sheets2)
            - [2.1.3.2. Select the next sheet in your workbook](#2132-select-the-next-sheet-in-your-workbook)
            - [2.1.3.3. Select the previous sheet in your workbook](#2133-select-the-previous-sheet-in-your-workbook)
        - [2.1.4. Selecting Worksheets](#214-selecting-worksheets)
            - [2.1.4.1. Select the `Task 2 Min and Max` tab](#2141-select-the-task-2-min-and-max-tab)
        - [2.1.5. s **2.1.4. Copying / Pasting Things**](#215-s-214-copying--pasting-things)
            - [2.1.5.1. Copy the value in cell `M3`](#2151-copy-the-value-in-cell-m3)
            - [2.1.5.2. Assign the cell `M2` the value `10`. Assign the cell `M3` the formula `=SUM(M2,2)`, which should equal `12`. Copy this formula. Paste this formula from `M4` down to `M100`.](#2152-assign-the-cell-m2-the-value-10-assign-the-cell-m3-the-formula-summ22-which-should-equal-12-copy-this-formula-paste-this-formula-from-m4-down-to-m100)
    - [2.2. **2.2. Formulas**](#22-22-formulas)
            - [2.2.0.3. Assigns each cell from `D40` to `F40` the formula found in cell `F28`](#2203-assigns-each-cell-from-d40-to-f40-the-formula-found-in-cell-f28)
            - [2.2.0.4. Assign each cell from `D40` to `F40` the formula found in cell `F28`](#2204-assign-each-cell-from-d40-to-f40-the-formula-found-in-cell-f28)
    - [2.3. **2.3. Functions**](#23-23-functions)
        - [2.3.1. **2.3.1. Math Functions**](#231-231-math-functions)
            - [2.3.1.1. Find the AVERAGE value of the range from `A2` to `A26` and place it in cell `A28`](#2311-find-the-average-value-of-the-range-from-a2-to-a26-and-place-it-in-cell-a28)
            - [2.3.1.2. Find the minimum of column `A`](#2312-find-the-minimum-of-column-a)
            - [2.3.1.3. Find the max of row `3`](#2313-find-the-max-of-row-3)
            - [2.3.1.4. Select all the values starting at `A2` down (`CTRL + DOWN`) in Excel, then sums the values](#2314-select-all-the-values-starting-at-a2-down-ctrl--down-in-excel-then-sums-the-values)
    - [2.4. **2.4. User Interaction**](#24-24-user-interaction)
        - [2.4.1. **2.4.1. Message Boxes**](#241-241-message-boxes)
            - [2.4.1.1. A simple message box with the text "Learning is kewl! and "OK" Button](#2411-a-simple-message-box-with-the-text-learning-is-kewl-and-ok-button)

<!-- /TOC -->

# 2. **2. Basic Functions**

### 2.0.1. Other

#### 2.0.1.1. Assigning Value to Cell ([Home](#1-table-of-contents))
`Range("M2").Value = 10`


Note here that when you are entering in a number, you do not need to put in quotation marks! If you were to write this instead:

`Range("M2").Value = "10"`

You would get a string value, not a integer (number) value.


#### 2.0.1.2. Assign the value "Yu Chen" to a range of cells from `A1` to `D2` ([Home](#1-table-of-contents))
`Range("A1:D2").Value = "Yu Chen"`

#### 2.0.1.3. Assign a value of `"Yu Chen"` to the variable `MyVariable`, and then assign this variable to cell `B2` ([Home](#1-table-of-contents))
```
MyVariable = "Yu Chen"
Range("B2").Value = MyVariable
```
([Home](#1-table-of-contents))
#### 2.0.1.4. Assigning Formula to Cell
This assigns the Excel formula to `M3` (take the value of `M2` and add `2` to it.)

`Range("M3").Formula = =SUM(M2,2)`

([Home](#1-table-of-contents))
## 2.1. **2.1. Selecting Things**

([Home](#1-table-of-contents))
### 2.1.1. **Selecting Workbooks**

([Home](#1-table-of-contents))
#### 2.1.1.1. Activating the current workbook (where the code resides)
`ThisWorkbook.Activate`

([Home](#1-table-of-contents))
#### 2.1.1.2. Activate a workbook to the name of `My Macro Book`
`Workbooks("My Work Book").Activate`

([Home](#1-table-of-contents))
#### 2.1.1.3. Activate the 2nd workbook (or specifically, the workbook in index position `2`)
`Workbooks(2).Activate`

([Home](#1-table-of-contents))
### 2.1.2. **2.1.1. Selecting Cells**

([Home](#1-table-of-contents))
#### 2.1.2.1. Select a single cell `A2`
`Range("A2").Select`

([Home](#1-table-of-contents))
#### 2.1.2.2. Select the cells from `A2` to `B10`

`Range("A2", "B10").Select`

([Home](#1-table-of-contents))
#### 2.1.2.3. Select the last row in column `A` of the dataset, and then moves `6` rows down 
`Range("A1").End(xlDown).Offset(6, 0).Select`

([Home](#1-table-of-contents))
#### 2.1.2.4. Select the last row in column `C` of the dataset, and then moves `6` rows down 
`Range("C1").End(xlDown).Offset(-6, 0).Select`

([Home](#1-table-of-contents))
#### 2.1.2.5. Select the **entire region of cells** 
This is equivalent to hitting `CTRL + SHIFT + DOWN + RIGHT` on your keyboard:

`ActiveCell.CurrentRegion.Select`

This is also equivalent to the following command:

`Range("A2", Range("A2").End(xlDown).End(xlToRight)).Select`

([Home](#1-table-of-contents))
#### 2.1.2.6. Select entire column that `A2` is on (column `A`)

`Range("A2").EntireColumn.Select`

([Home](#1-table-of-contents))
#### 2.1.2.7. Select entire row that `A2` is on (row `2`) 
`Range("A2").EntireRow.Select`

([Home](#1-table-of-contents))
### 2.1.3. **2.1.2. Selecting Sheets**

A `Sheet` and a `Worksheet` are related, but cannot be used interchangeably. A `Sheet` is any Excel sheet, whereas a `Worksheet` is only a regular Excel worksheet. For example, a chart is a `Sheet` but is not a `Worksheet`.

([Home](#1-table-of-contents))
#### 2.1.3.1. Select a sheet by tab name (`Sheets2`)
`Sheets("Sheets2").Select`

([Home](#1-table-of-contents))
#### 2.1.3.2. Select the next sheet in your workbook

`ActiveSheet.Next.Select`

([Home](#1-table-of-contents))
#### 2.1.3.3. Select the previous sheet in your workbook

`ActiveSheet.Previous.Select`

([Home](#1-table-of-contents))
### 2.1.4. Selecting Worksheets


#### 2.1.4.1. Select the `Task 2 Min and Max` tab 
`Worksheets("Task 2 Min and Max").Activate`


### 2.1.5. s **2.1.4. Copying / Pasting Things**

([Home](#1-table-of-contents))
#### 2.1.5.1. Copy the value in cell `M3`
`Range("M3").Copy`

([Home](#1-table-of-contents))
#### 2.1.5.2. Assign the cell `M2` the value `10`. Assign the cell `M3` the formula `=SUM(M2,2)`, which should equal `12`. Copy this formula. Paste this formula from `M4` down to `M100`.

```
Range("M2").Value = 10 
Range("M3").Formula = "=SUM(M2,2)"
Range("M3").Copy 
Range("M4:D100").PasteSpecial
```

## 2.2. **2.2. Formulas**

#### 2.2.0.3. Assigns each cell from `D40` to `F40` the formula found in cell `F28` 
`Range("D40:F40").Formula = Range("F28").Formula`

([Home](#1-table-of-contents))
#### 2.2.0.4. Assign each cell from `D40` to `F40` the formula found in cell `F28`
`Range("D40:F40").Formula = Range("F28").Formula`

## 2.3. **2.3. Functions**

### 2.3.1. **2.3.1. Math Functions**

([Home](#1-table-of-contents))
#### 2.3.1.1. Find the AVERAGE value of the range from `A2` to `A26` and place it in cell `A28` 
`Range("A28").Value = Application.WorksheetFunction.Average(Range("A2:A26"))`

([Home](#1-table-of-contents))
#### 2.3.1.2. Find the minimum of column `A`

`Application.WorksheetFunction.Min(Range("A2").EntireColumn.Select)`

You need to input a range inside the `Min()` brackets.

([Home](#1-table-of-contents))
#### 2.3.1.3. Find the max of row `3`
`Application.WorksheetFunction.Min(Range("A3").EntireRow.Select)`

([Home](#1-table-of-contents))
#### 2.3.1.4. Select all the values starting at `A2` down (`CTRL + DOWN`) in Excel, then sums the values 
`Range("A30").Value = Application.WorksheetFunction.Sum(Range("A2", Range("A2").End(xlDown)))`


## 2.4. **2.4. User Interaction**

### 2.4.1. **2.4.1. Message Boxes**


#### 2.4.1.1. A simple message box with the text "Learning is kewl! and "OK" Button

`MsgBox ("Learning is kewl!")`

([Home](#1-table-of-contents))