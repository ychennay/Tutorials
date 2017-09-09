# 1. Table of Contents

<!-- TOC -->

- [1. Table of Contents](#1-table-of-contents)
- [2. **Formatting**](#2-formatting)
        - [**Color**](#color)
        - [**Font**](#font)
            - [Set the font in `A9` to bold](#set-the-font-in-a9-to-bold)
            - [Set the font in `A2` to be regular](#set-the-font-in-a2-to-be-regular)
            - [Set the cell `B4` to be both bold and italic:](#set-the-cell-b4-to-be-both-bold-and-italic)
- [3. **Basic Functions**](#3-basic-functions)
        - [3.0.1. Other](#301-other)
            - [3.0.1.1. Assigning Value to Cell](#3011-assigning-value-to-cell)
            - [3.0.1.2. Assign the value "Yu Chen" to a range of cells from `A1` to `D2`](#3012-assign-the-value-yu-chen-to-a-range-of-cells-from-a1-to-d2)
            - [3.0.1.3. Assign a value of `"Yu Chen"` to the variable `MyVariable`, and then assign this variable to cell `B2`](#3013-assign-a-value-of-yu-chen-to-the-variable-myvariable-and-then-assign-this-variable-to-cell-b2)
            - [3.0.1.4. Assigning Formula to Cell](#3014-assigning-formula-to-cell)
    - [3.1. **2.1. Selecting Things**](#31-21-selecting-things)
        - [3.1.1. **Selecting Workbooks**](#311-selecting-workbooks)
            - [3.1.1.1. **Activating the current workbook (where the code resides)**](#3111-activating-the-current-workbook-where-the-code-resides)
            - [3.1.1.2. **Activate a workbook to the name of `My Macro Book`**](#3112-activate-a-workbook-to-the-name-of-my-macro-book)
            - [3.1.1.3. **Activate the 2nd workbook (or specifically, the workbook in index position `2`)**](#3113-activate-the-2nd-workbook-or-specifically-the-workbook-in-index-position-2)
        - [3.1.2. **Selecting Cells**](#312-selecting-cells)
            - [3.1.2.1. **Select a single cell `A2`**](#3121-select-a-single-cell-a2)
            - [3.1.2.2. **Select the cells from `A2` to `B10`**](#3122-select-the-cells-from-a2-to-b10)
            - [3.1.2.3. **Select the last row in column `A` of the dataset, and then moves `6` rows down**](#3123-select-the-last-row-in-column-a-of-the-dataset-and-then-moves-6-rows-down)
            - [3.1.2.4. **Select the last row in column `C` of the dataset, and then moves `6` rows down**](#3124-select-the-last-row-in-column-c-of-the-dataset-and-then-moves-6-rows-down)
            - [3.1.2.5. **Select the entire region of cells**](#3125-select-the-entire-region-of-cells)
            - [3.1.2.6. **Select entire column that `A2` is on (column `A`)**](#3126-select-entire-column-that-a2-is-on-column-a)
            - [3.1.2.7. Select entire row that `A2` is on (row `2`)](#3127-select-entire-row-that-a2-is-on-row-2)
        - [3.1.3. **Selecting Sheets**](#313-selecting-sheets)
            - [3.1.3.1. Select a sheet by tab name (`Sheets2`)**](#3131-select-a-sheet-by-tab-name-sheets2)
            - [3.1.3.2. ** Select the next sheet in your workbook**](#3132--select-the-next-sheet-in-your-workbook)
            - [3.1.3.3. **Select the previous sheet in your workbook**](#3133-select-the-previous-sheet-in-your-workbook)
        - [3.1.4. Selecting Worksheets](#314-selecting-worksheets)
            - [3.1.4.1. **Select the `Task 2 Min and Max` tab**](#3141-select-the-task-2-min-and-max-tab)
        - [3.1.5. s **2.1.4. Copying / Pasting Things**](#315-s-214-copying--pasting-things)
            - [3.1.5.1. Copy the value in cell `M3`](#3151-copy-the-value-in-cell-m3)
            - [3.1.5.2. Copy and paste value from cell `A1` to `B1` all in one line](#3152-copy-and-paste-value-from-cell-a1-to-b1-all-in-one-line)
            - [3.1.5.3. Assign the cell `M2` the value `10`. Assign the cell `M3` the formula `=SUM(M2,2)`, which should equal `12`. Copy this formula. Paste this formula from `M4` down to `M100`.](#3153-assign-the-cell-m2-the-value-10-assign-the-cell-m3-the-formula-summ22-which-should-equal-12-copy-this-formula-paste-this-formula-from-m4-down-to-m100)
    - [3.2. **Formulas**](#32-formulas)
            - [3.2.0.4. Assigns each cell from `D40` to `F40` the formula found in cell `F28`](#3204-assigns-each-cell-from-d40-to-f40-the-formula-found-in-cell-f28)
            - [3.2.0.5. Assign each cell from `D40` to `F40` the formula found in cell `F28`](#3205-assign-each-cell-from-d40-to-f40-the-formula-found-in-cell-f28)
    - [3.3. **2.3. Functions**](#33-23-functions)
        - [3.3.1. **Math Functions**](#331-math-functions)
            - [3.3.1.1. Find the AVERAGE value of the range from `A2` to `A26` and place it in cell `A28`](#3311-find-the-average-value-of-the-range-from-a2-to-a26-and-place-it-in-cell-a28)
            - [3.3.1.2. Find the minimum of column `A`](#3312-find-the-minimum-of-column-a)
            - [3.3.1.3. Find the max of row `3`](#3313-find-the-max-of-row-3)
            - [3.3.1.4. Select all the values starting at `A2` down (`CTRL + DOWN`) in Excel, then sums the values](#3314-select-all-the-values-starting-at-a2-down-ctrl--down-in-excel-then-sums-the-values)
    - [3.4. **User Interaction**](#34-user-interaction)
        - [3.4.1. **Message Boxes**](#341-message-boxes)
            - [3.4.1.1. A simple message box with the text "Learning is kewl! and "OK" Button](#3411-a-simple-message-box-with-the-text-learning-is-kewl-and-ok-button)

<!-- /TOC -->

# 2. **Formatting**

### **Color**

### **Font**

#### Set the font in `A9` to bold
`Range("A9").Font.Bold = True`

You can also set it this way:

`Range("A9").Font.FontStyle = "Bold"`

#### Set the font in `A2` to be regular
`Range("A9").Font.Bold = False`

You can also set it this way:

`Range("A9").Font.FontStyle = "Regular"`

#### Set the cell `B4` to be both bold and italic:

`Range("B4").Font.FontStyle = "Bold italic"`


# 3. **Basic Functions**

### 3.0.1. Other

#### 3.0.1.1. Assigning Value to Cell
`Range("M2").Value = 10`


Note here that when you are entering in a number, you do not need to put in quotation marks! If you were to write this instead:

`Range("M2").Value = "10"`

You would get a string value, not a integer (number) value.


#### 3.0.1.2. Assign the value "Yu Chen" to a range of cells from `A1` to `D2`
`Range("A1:D2").Value = "Yu Chen"`

#### 3.0.1.3. Assign a value of `"Yu Chen"` to the variable `MyVariable`, and then assign this variable to cell `B2`
```
MyVariable = "Yu Chen"
Range("B2").Value = MyVariable
```
([Home](#1-table-of-contents))
#### 3.0.1.4. Assigning Formula to Cell
This assigns the Excel formula to `M3` (take the value of `M2` and add `2` to it.)

`Range("M3").Formula = =SUM(M2,2)`

([Home](#1-table-of-contents))
## 3.1. **2.1. Selecting Things**

([Home](#1-table-of-contents))
### 3.1.1. **Selecting Workbooks**

([Home](#1-table-of-contents))
#### 3.1.1.1. **Activating the current workbook (where the code resides)**
`ThisWorkbook.Activate`

([Home](#1-table-of-contents))
#### 3.1.1.2. **Activate a workbook to the name of `My Macro Book`**
`Workbooks("My Work Book").Activate`

([Home](#1-table-of-contents))
#### 3.1.1.3. **Activate the 2nd workbook (or specifically, the workbook in index position `2`)**
`Workbooks(2).Activate`

([Home](#1-table-of-contents))
### 3.1.2. **Selecting Cells**

([Home](#1-table-of-contents))
#### 3.1.2.1. **Select a single cell `A2`**
`Range("A2").Select`

([Home](#1-table-of-contents))
#### 3.1.2.2. **Select the cells from `A2` to `B10`**

`Range("A2", "B10").Select`

([Home](#1-table-of-contents))
#### 3.1.2.3. **Select the last row in column `A` of the dataset, and then moves `6` rows down** 
`Range("A1").End(xlDown).Offset(6, 0).Select`

([Home](#1-table-of-contents))
#### 3.1.2.4. **Select the last row in column `C` of the dataset, and then moves `6` rows down** 
`Range("C1").End(xlDown).Offset(-6, 0).Select`

([Home](#1-table-of-contents))
#### 3.1.2.5. **Select the entire region of cells** 
This is equivalent to hitting `CTRL + SHIFT + DOWN + RIGHT` on your keyboard:

`ActiveCell.CurrentRegion.Select`

This is also equivalent to the following command:

`Range("A2", Range("A2").End(xlDown).End(xlToRight)).Select`

([Home](#1-table-of-contents))
#### 3.1.2.6. **Select entire column that `A2` is on (column `A`)**

`Range("A2").EntireColumn.Select`

([Home](#1-table-of-contents))
#### 3.1.2.7. Select entire row that `A2` is on (row `2`) 
`Range("A2").EntireRow.Select`

([Home](#1-table-of-contents))
### 3.1.3. **Selecting Sheets**

A `Sheet` and a `Worksheet` are related, but cannot be used interchangeably. A `Sheet` is any Excel sheet, whereas a `Worksheet` is only a regular Excel worksheet. For example, a chart is a `Sheet` but is not a `Worksheet`.

([Home](#1-table-of-contents))
#### 3.1.3.1. Select a sheet by tab name (`Sheets2`)**
`Sheets("Sheets2").Select`

([Home](#1-table-of-contents))
#### 3.1.3.2. ** Select the next sheet in your workbook**

`ActiveSheet.Next.Select`

([Home](#1-table-of-contents))
#### 3.1.3.3. **Select the previous sheet in your workbook**

`ActiveSheet.Previous.Select`

([Home](#1-table-of-contents))
### 3.1.4. Selecting Worksheets

#### 3.1.4.1. **Select the `Task 2 Min and Max` tab**
`Worksheets("Task 2 Min and Max").Activate`

([Home](#1-table-of-contents))

### 3.1.5. s **2.1.4. Copying / Pasting Things**

([Home](#1-table-of-contents))
#### 3.1.5.1. Copy the value in cell `M3`
`Range("M3").Copy`

#### 3.1.5.2. Copy and paste value from cell `A1` to `B1` all in one line
`Range("A1").Copy Range("B1")`

Note that this pastes all the formatting as well, so if you had a bolded cell in `A1`, you'll also have a bolded cell in `B1`.

([Home](#1-table-of-contents))
#### 3.1.5.3. Assign the cell `M2` the value `10`. Assign the cell `M3` the formula `=SUM(M2,2)`, which should equal `12`. Copy this formula. Paste this formula from `M4` down to `M100`.

```
Range("M2").Value = 10 
Range("M3").Formula = "=SUM(M2,2)"
Range("M3").Copy 
Range("M4:D100").PasteSpecial
```
([Home](#1-table-of-contents))

## 3.2. **Formulas**

#### 3.2.0.4. Assigns each cell from `D40` to `F40` the formula found in cell `F28` 
`Range("D40:F40").Formula = Range("F28").Formula`

([Home](#1-table-of-contents))
#### 3.2.0.5. Assign each cell from `D40` to `F40` the formula found in cell `F28`
`Range("D40:F40").Formula = Range("F28").Formula`

([Home](#1-table-of-contents))

## 3.3. **2.3. Functions**

### 3.3.1. **Math Functions**

([Home](#1-table-of-contents))
#### 3.3.1.1. Find the AVERAGE value of the range from `A2` to `A26` and place it in cell `A28` 
`Range("A28").Value = Application.WorksheetFunction.Average(Range("A2:A26"))`

([Home](#1-table-of-contents))
#### 3.3.1.2. Find the minimum of column `A`

`Application.WorksheetFunction.Min(Range("A2").EntireColumn.Select)`

You need to input a range inside the `Min()` brackets.

([Home](#1-table-of-contents))
#### 3.3.1.3. Find the max of row `3`
`Application.WorksheetFunction.Min(Range("A3").EntireRow.Select)`

([Home](#1-table-of-contents))
#### 3.3.1.4. Select all the values starting at `A2` down (`CTRL + DOWN`) in Excel, then sums the values 
`Range("A30").Value = Application.WorksheetFunction.Sum(Range("A2", Range("A2").End(xlDown)))`

([Home](#1-table-of-contents))
## 3.4. **User Interaction**

### 3.4.1. **Message Boxes**


#### 3.4.1.1. A simple message box with the text "Learning is kewl! and "OK" Button

`MsgBox ("Learning is kewl!")`

([Home](#1-table-of-contents))