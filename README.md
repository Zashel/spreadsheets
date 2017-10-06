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

As with a real spreadsheet you can get the rows you want...

```
>> spreadsheet.Rows[3:3]
[[7, 8, 9]]
>> spreadsheet.Rows[4:5]
[[10, 11, 12], [13, 14, 15]]
```

and the columns you want too.

```
>> spreadsheet.Columns[1:2]
[["Column2", 2, 5, 8, 11, 14], ["Column3", 3, 6, 9, 12, 15]]
```

You can modify Row or a Column and the spreadsheet will be modified.

```
>> spreadsheet.Columns[1:2][2:0] = 22 # It was 5 before
>> spreadsheet
[["Column1", "Column2", "Column3"], [1, 2, 3], [4, 22, 6], [7, 8, 9], [10, 11, 12], [13, 14, 15]]
>> spreadsheet[1:3]
22
```

You can assign a Cell not just the value, but another cell.

```
>> spreadsheet[1:3] = spreadsheet[1:4] # We change the last 22 to the 8 beneath
>> spreadsheet
[["Column1", "Column2", "Column3"], [1, 2, 3], [4, 8, 6], [7, 8, 9], [10, 11, 12], [13, 14, 15]]
>> spreadsheet[1:4] = 16 # We now change the original value of 8 changing both
>> spreadsheet
[["Column1", "Column2", "Column3"], [1, 2, 3], [4, 16, 6], [7, 16, 9], [10, 11, 12], [13, 14, 15]]
```

### Considerations

The search by [col:row] may be very difficult to aprehend, so it accepts [row][col] too.
