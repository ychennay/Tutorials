# 1. Table of Contents

<!-- TOC -->

- [1. Table of Contents](#1-table-of-contents)
- [2. Basic Functions](#2-basic-functions)
            - [2.0.0.1. Assigning Value to Cell](#2001-assigning-value-to-cell)
            - [2.0.0.2. Assigning Formula to Cell](#2002-assigning-formula-to-cell)
    - [2.1. Selecting Things](#21-selecting-things)
        - [2.1.1. Selecting Cells](#211-selecting-cells)
            - [2.1.1.1. Selects a single cell `A2`](#2111-selects-a-single-cell-a2)
            - [2.1.1.2. Selects the cells from `A2` to `B10`](#2112-selects-the-cells-from-a2-to-b10)
            - [2.1.1.3. Select the **entire region of cells**](#2113-select-the-entire-region-of-cells)
            - [2.1.1.4. Select entire column that `A2` is on (column `A`)](#2114-select-entire-column-that-a2-is-on-column-a)
            - [2.1.1.5. Select entire row that `A2` is on (row `2`)](#2115-select-entire-row-that-a2-is-on-row-2)
        - [2.1.2. Selecting Sheets](#212-selecting-sheets)
            - [2.1.2.1. Select the next sheet in your workbook](#2121-select-the-next-sheet-in-your-workbook)
            - [2.1.2.2. Select the previous sheet in your workbook](#2122-select-the-previous-sheet-in-your-workbook)
        - [Selecting Worksheets](#selecting-worksheets)
            - [Select the `Task 2 Min and Max` tab](#select-the-task-2-min-and-max-tab)
        - [2.1.3. Copying / Pasting Things](#213-copying--pasting-things)
            - [2.1.3.1. Copy the value in cell `M3`](#2131-copy-the-value-in-cell-m3)
            - [Assign the cell `M2` the value `10`. Assign the cell `M3` the formula `=SUM(M2,2)`, which should equal `12`. Copy this formula. Paste this formula from `M4` down to `M100`.](#assign-the-cell-m2-the-value-10-assign-the-cell-m3-the-formula-summ22-which-should-equal-12-copy-this-formula-paste-this-formula-from-m4-down-to-m100)
    - [Formulas](#formulas)
            - [Assigns each cell from `D40` to `F40` the formula found in cell `F28`](#assigns-each-cell-from-d40-to-f40-the-formula-found-in-cell-f28)
    - [2.2. Functions](#22-functions)
        - [2.2.1. Math Functions](#221-math-functions)
            - [2.2.1.1. Find the minimum of column `A`](#2211-find-the-minimum-of-column-a)
            - [2.2.1.2. Find the max of row `3`](#2212-find-the-max-of-row-3)
    - [2.3. User Interaction](#23-user-interaction)
        - [2.3.1. Message Boxes](#231-message-boxes)
            - [2.3.1.1. A simple message box with the text "Learning is kewl! and "OK" Button](#2311-a-simple-message-box-with-the-text-learning-is-kewl-and-ok-button)

<!-- /TOC -->

# 2. Basic Functions

#### 2.0.0.1. Assigning Value to Cell
`Range("M2").Value = 10`

#### 2.0.0.2. Assigning Formula to Cell
This assigns the Excel formula to `M3` (take the value of `M2` and add `2` to it.)

`Range("M3").Formula = =SUM(M2,2)`

## 2.1. Selecting Things

### 2.1.1. Selecting Cells

#### 2.1.1.1. Selects a single cell `A2`
`Range("A2").Select`

#### 2.1.1.2. Selects the cells from `A2` to `B10`

`Range("A2", "B10").Select`

#### 2.1.1.3. Select the **entire region of cells** 
This is equivalent to hitting `CTRL + SHIFT + DOWN + RIGHT` on your keyboard:

`ActiveCell.CurrentRegion.Select`

This is also equivalent to the following command:

`Range("A2", Range("A2").End(xlDown).End(xlToRight)).Select`

#### 2.1.1.4. Select entire column that `A2` is on (column `A`)

`Range("A2").EntireColumn.Select`

#### 2.1.1.5. Select entire row that `A2` is on (row `2`) 
`Range("A2").EntireRow.Select`

### 2.1.2. Selecting Sheets

#### 2.1.2.1. Select the next sheet in your workbook

`ActiveSheet.Next.Select`

#### 2.1.2.2. Select the previous sheet in your workbook

`ActiveSheet.Previous.Select`

### Selecting Worksheets

#### Select the `Task 2 Min and Max` tab 
`Worksheets("Task 2 Min and Max").Activate`

### 2.1.3. Copying / Pasting Things

#### 2.1.3.1. Copy the value in cell `M3`
`Range("M3").Copy`

#### Assign the cell `M2` the value `10`. Assign the cell `M3` the formula `=SUM(M2,2)`, which should equal `12`. Copy this formula. Paste this formula from `M4` down to `M100`.

```
Range("M2").Value = 10 
Range("M3").Formula = "=SUM(M2,2)"
Range("M3").Copy 
Range("M4:D100").PasteSpecial
```

## Formulas

#### Assigns each cell from `D40` to `F40` the formula found in cell `F28` 
`Range("D40:F40").Formula = Range("F28").Formula`

## 2.2. Functions

### 2.2.1. Math Functions

#### 2.2.1.1. Find the minimum of column `A`

`Application.WorksheetFunction.Min(Range("A2").EntireColumn.Select)`

You need to input a range inside the `Min()` brackets.

#### 2.2.1.2. Find the max of row `3`
`Application.WorksheetFunction.Min(Range("A3").EntireRow.Select)`

## 2.3. User Interaction

### 2.3.1. Message Boxes

#### 2.3.1.1. A simple message box with the text "Learning is kewl! and "OK" Button

`MsgBox ("Learning is kewl!")`
