# zashel.spreadsheets
Spreadsheet-like in Python.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

First of all, it's imperative not to depend on other libraries but those builtins.

### First use

We can create the spreadsheet just instantiating

```
>> spreadsheet = Spreadsheet(["Column1", "Column2", "Column3"],
                             [1, 2, 3],
                             [4, 5, 6])
>> spreadsheet
[["Column1", "Column2", "Column3"], [1, 2, 3], [4, 5, 6]]
```

It has all the lists methods, so you can append, extend... sort is a future feature.

```
>> spreadsheet.append([7, 8, 9])
>> spreadsheet
[["Column1", "Column2", "Column3"], [1, 2, 3], [4, 5, 6], [7, 8, 9]]
>> spreadsheet.extend([[10, 11, 12], [13, 14, 15]])
>> spreadsheet
[["Column1", "Column2", "Column3"], [1, 2, 3], [4, 5, 6], [7, 8, 9], [10, 11, 12], [13, 14, 15]]
```

You can call to a specific cell by its position by [col:row]. All coordinates begin with 0.

```
>> spreadsheet[2:2]
6
```

You can call by typical spreadsheet nomenclature:

```
>> spreadsheet["C3"]
6
```

As with a real spreadsheet you can get the rows you want...

```
>> spreadsheet.Rows[3:3]
[[7, 8, 9]]
>> spreadsheet["4:5"] #This works as before
[[10, 11, 12], [13, 14, 15]]
```

and the columns you want too.

```
>> spreadsheet.Columns[1:2]
[["Column2", 2, 5, 8, 11, 14], ["Column3", 3, 6, 9, 12, 15]]
>> spreadsheet["B:C"]
[["Column2", 2, 5, 8, 11, 14], ["Column3", 3, 6, 9, 12, 15]]
```

You can modify Row or a Column and the spreadsheet will be modified.

```
>> spreadsheet["B:C"][2:0] = 22 # It was 5 before
>> spreadsheet
[["Column1", "Column2", "Column3"], [1, 2, 3], [4, 22, 6], [7, 8, 9], [10, 11, 12], [13, 14, 15]]
>> spreadsheet["B3"]
22
```

You can assign a Cell not just the value, but another cell.

```
>> spreadsheet["B3"] = spreadsheet["B4"] # We change the last 22 to the 8 beneath.
>> spreadsheet["B3"] = "=B4" # This is accepted too 
>> spreadsheet
[["Column1", "Column2", "Column3"], [1, 2, 3], [4, 8, 6], [7, 8, 9], [10, 11, 12], [13, 14, 15]]
>> spreadsheet["B4"] = 16 # We now change the original value of 8 changing both
>> spreadsheet
[["Column1", "Column2", "Column3"], [1, 2, 3], [4, 16, 6], [7, 16, 9], [10, 11, 12], [13, 14, 15]]
```

You can assign a function to a cell, as with a real spreadsheet.

```
>> spreadsheet.append(["=sum(A2:A6)", "=sum(B2:B6)", "=sum(C2:C6)"])
>> spreadsheet
[["Column1", "Column2", "Column3"], [1, 2, 3], [4, 8, 6], [7, 8, 9], [10, 11, 12], [13, 14, 15], [35, 59, 45]]
>> spreadsheet["A2"] = 2
>> spreadsheet
[["Column1", "Column2", "Column3"], [2, 2, 3], [4, 8, 6], [7, 8, 9], [10, 11, 12], [13, 14, 15], [36, 59, 45]]
```

A minimal set of functions have already been declared: average, count, max, min and sum