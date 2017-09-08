# 1. Table of Contents

<!-- TOC -->

- [1. Table of Contents](#1-table-of-contents)
- [2. Basic Functions](#2-basic-functions)
    - [2.1. Selecting Things](#21-selecting-things)
        - [2.1.1. Selecting Cells](#211-selecting-cells)
            - [2.1.1.1. Selects a single cell `A2`](#2111-selects-a-single-cell-a2)
            - [2.1.1.2. Selects the cells from `A2` to `B10`](#2112-selects-the-cells-from-a2-to-b10)
            - [2.1.1.3. Select the **entire region of cells**](#2113-select-the-entire-region-of-cells)
            - [2.1.1.4. Select entire column that `A2` is on (column A)](#2114-select-entire-column-that-a2-is-on-column-a)
            - [2.1.1.5. Select entire row that A2 is on (row 2) Range("A2").EntireRow.Select](#2115-select-entire-row-that-a2-is-on-row-2-rangea2entirerowselect)
        - [2.1.2. Selecting Sheets](#212-selecting-sheets)
            - [2.1.2.1. Select the next sheet in your workbook](#2121-select-the-next-sheet-in-your-workbook)
            - [2.1.2.2. Select the previous sheet in your workbook](#2122-select-the-previous-sheet-in-your-workbook)
    - [2.2. User Interaction](#22-user-interaction)
        - [2.2.1. Message Boxes](#221-message-boxes)

<!-- /TOC -->

# 2. Basic Functions

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

#### 2.1.1.4. Select entire column that `A2` is on (column A)

`Range("A2").EntireColumn.Select`

#### 2.1.1.5. Select entire row that A2 is on (row 2) Range("A2").EntireRow.Select

### 2.1.2. Selecting Sheets

#### 2.1.2.1. Select the next sheet in your workbook

`ActiveSheet.Next.Select`

#### 2.1.2.2. Select the previous sheet in your workbook

`ActiveSheet.Previous.Select`


## 2.2. User Interaction

### 2.2.1. Message Boxes

A message box will pop up with the text "Learning is kewl!"

`MsgBox ("Learning is kewl!")`
