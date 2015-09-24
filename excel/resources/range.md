# Range Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Range represents a set of one or more contiguous cells such as a cell, a row, a column, block of cells, etc.

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|address|string|Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. Sheet1!A1:B4). Read-only.|
|addressLocal|string|Represents range reference for the specified range in the language of the user. Read-only.|
|cellCount|int|Number of cells in the range. Read-only.|
|columnCount|int|Represents the total number of columns in the range. Read-only.|
|columnIndex|int|Represents the column number of the first cell in the range. Zero-indexed. Read-only.|
|columnWidth|number|Returns or sets the width of all columns in the specified range, measured in points. |
|formulas|object[][]|Represents the formula in A1-style notation.|
|formulasLocal|object[][]|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|
|formulasR1C1|string|Returns or sets the formula for the object, using R1C1-style notation.|
|numberFormat|object[][]|Represents Excel's number format code for the given cell.|
|orientation|number|Returns or sets a integer value that represents the text orientation, from -90 to 90 degrees. |
|rowCount|int|Returns the total number of rows in the range. Read-only.|
|rowHeight|number|Returns or sets the height of the first row in the range specified, measured in points.|
|rowIndex|int|Returns the row number of the first cell in the range. Zero-indexed. Read-only.|
|text|object[][]|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|
|valueTypes|string|Represents the type of data of each cell. Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|
|values|object[][]|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|format|[RangeFormat](rangeformat.md)|Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties. Read-only.|
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current range. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[autoFitColumn()](#autofitcolumn)|[Range](range.md)|Changes the width of the columns in the range to achieve the best fit.|
|[autoFitRow()](#autofitrow)|[Range](range.md)|Changes the height of the rows in the range to achieve the best fit.|
|[clear(applyTo: string)](#clearapplyto-string)|void|Clear range values, format, fill, border, etc.|
|[delete(shift: string)](#deleteshift-string)|void|Deletes the cells associated with the range.|
|[getBoundingRect(anotherRange: Range or string)](#getboundingrectanotherrange-range-or-string)|[Range](range.md)|Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E16".|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.|
|[getColumn(column: number)](#getcolumncolumn-number)|[Range](range.md)|Gets a column contained in the range.|
|[getEntireColumn()](#getentirecolumn)|[Range](range.md)|Gets an object that represents the entire column of the range.|
|[getEntireRow()](#getentirerow)|[Range](range.md)|Gets an object that represents the entire row of the range.|
|[getIntersection(anotherRange: Range or string)](#getintersectionanotherrange-range-or-string)|[Range](range.md)|Gets the range object that represents the rectangular intersection of the given ranges.|
|[getLastCell()](#getlastcell)|[Range](range.md)|Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".|
|[getLastColumn()](#getlastcolumn)|[Range](range.md)|Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".|
|[getLastRow()](#getlastrow)|[Range](range.md)|Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[Range](range.md)|Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an exception will be thrown.|
|[getRow(row: number)](#getrowrow-number)|[Range](range.md)|Gets a row contained in the range.|
|[getUsedRange(valuesOnly?:boolean)](#getusedrangevaluesonlyboolean)|[Range](range.md)|Returns the used range of the given range object.|
|[insert(shift: string)](#insertshift-string)|[Range](range.md)|Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[merge(across: boolean)](#mergeacross-boolean)|[Range](range.md)|Creates a merged cell from the specified Range object.|
|[select()](#select)|void|Selects the specified range in the Excel UI.|
|[unmerge()](#unmerge)|[Range](range.md)|Separates a merged area into individual cells. If the range contains no merged cell, this API will do no operations on the Range.|

## Method Details

### autoFitColumn()
Changes the width of the columns in the range to achieve the best fit.

#### Syntax
```js
rangeObject.autoFitColumn();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js

```
### autoFitRow()
Changes the height of the rows in the range to achieve the best fit.

#### Syntax
```js
rangeObject.autoFitRow();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js

```
### clear(applyTo: string)
Clear range values, format, fill, border, etc.

#### Syntax
```js
rangeObject.clear(applyTo);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|applyTo|string|Optional. Determines the type of clear action. Possible values are: `All` Default-option,`Formats` ,`Contents` |

#### Returns
void

#### Examples

Below example clears format and contents of the range. 

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.clear();
	return ctx.sync(); 
}); 
```

### delete(shift: string)
Deletes the cells associated with the range.

#### Syntax
```js
rangeObject.delete(shift);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|shift|string|Specifies which way to shift the cells.  Possible values are: Up, Left|

#### Returns
void

#### Examples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.delete();
	return ctx.sync(); 
}); 
```

### getBoundingRect(anotherRange: Range or string)
Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E16".

#### Syntax
```js
rangeObject.getBoundingRect(anotherRange);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|anotherRange|Range or string|The range object or address or range name.|

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D4:G6";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var range = range.getBoundingRect("G4:H8");
	range.load(address);
	return ctx.sync().then(function() {
		Console.log(range.address); // Prints Sheet1!D4:H8
	});
});
```

### getCell(row: number, column: number)
Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.

#### Syntax
```js
rangeObject.getCell(row, column);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|row|number|Row number of the cell to be retrieved. Zero-indexed.|
|column|number|Column number of the cell to be retrieved. Zero-indexed.|

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var cell = range.cell(0,0);
	cell.load(address);
	return ctx.sync().then(function() {
		Console.log(cell.address);
	});
});
```

### getColumn(column: number)
Gets a column contained in the range.

#### Syntax
```js
rangeObject.getColumn(column);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|column|number|Column number of the range to be retrieved. Zero-indexed.|

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet19";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getColumn(1);
	range.load(address);
	return ctx.sync().then(function() {
		Console.log(range.address); // prints Sheet1!B1:B8
	});
});
```

### getEntireColumn()
Gets an object that represents the entire column of the range.

#### Syntax
```js
rangeObject.getEntireColumn();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

Note: the grid properties of the Range (values, numberFormat, formula) contains `null` since the Range in question is unbounded.

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var rangeEC = range.getEntireColumn();
	rangeEC.load(address);
	return ctx.sync().then(function() {
		Console.log(rangeEC.address);
	});
});
```
### getEntireRow()
Gets an object that represents the entire row of the range.

#### Syntax
```js
rangeObject.getEntireRow();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js

Excel.run(function (ctx) {
	var sheetName = "Sheet1";
	var rangeAddress = "D:F"; 
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var rangeER = range.getEntireRow();
	rangeER.load(address);
	return ctx.sync().then(function() {
		Console.log(rangeER.address);
	});
});
```
The grid properties of the Range (values, numberFormat, formula) contains `null` since the Range in question is unbounded.

### getIntersection(anotherRange: Range or string)
Gets the range object that represents the rectangular intersection of the given ranges.

#### Syntax
```js
rangeObject.getIntersection(anotherRange);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|anotherRange|Range or string|The range object or range address that will be used to determine the intersection of ranges.|

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getIntersection("D4:G6");
	range.load(address);
	return ctx.sync().then(function() {
		Console.log(range.address); // prints Sheet1!D4:F6
	});
});
```

### getLastCell()
Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".

#### Syntax
```js
rangeObject.getLastCell();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastCell();
	range.load(address);
	return ctx.sync().then(function() {
		Console.log(range.address); // prints Sheet1!F8
	});
});
```

### getLastColumn()
Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".

#### Syntax
```js
rangeObject.getLastColumn();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastColumn();
	range.load(address);
	return ctx.sync().then(function() {
		Console.log(range.address); // prints Sheet1!F1:F8
	});
});
```

### getLastRow()
Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".

#### Syntax
```js
rangeObject.getLastRow();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastRow();
	range.load(address);
	return ctx.sync().then(function() {
		Console.log(range.address); // prints Sheet1!A8:F8
	});
});
```


### getOffsetRange(rowOffset: number, columnOffset: number)
Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an exception will be thrown.

#### Syntax
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowOffset|number|The number of rows (positive, negative, or 0) by which the range is to be offset. Positive values are offset downward, and negative values are offset upward.|
|columnOffset|number|The number of columns (positive, negative, or 0) by which the range is to be offset. Positive values are offset to the right, and negative values are offset to the left.|

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D4:F6";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getOffsetRange(-1,4);
	range.load(address);
	return ctx.sync().then(function() {
		Console.log(range.address); // prints Sheet1!H3:K5
	});
});
```

### getRow(row: number)
Gets a row contained in the range.

#### Syntax
```js
rangeObject.getRow(row);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|row|number|Row number of the range to be retrieved. Zero-indexed.|

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getRow(1);
	range.load(address);
	return ctx.sync().then(function() {
		Console.log(range.address); // prints Sheet1!A2:F2
	});
});
```

### getUsedRange(valuesOnly?:boolean)
Returns the used range of the given range object.

#### Syntax
```js
rangeObject.getUsedRange(valuesOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|valuesOnly|boolean|Optional. True to get the used Range with values only. Default is false.|

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var rangeUR = range.getUsedRange();
	rangeUR.load(address);
	return ctx.sync().then(function() {
		Console.log(rangeUR.address);
	});
});
```

### insert(shift: string)
Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.

#### Syntax
```js
rangeObject.insert(shift);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|shift|string|Specifies which way to shift the cells.  Possible values are: Down, Right|

#### Returns
[Range](range.md)

#### Examples

```js
	
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F5:F10";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.insert();
	return ctx.sync(); 
	}); 
});
```

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

#### Examples
```js

```
### merge(across: boolean)
Creates a merged cell from the specified Range object.

#### Syntax
```js
rangeObject.merge(across);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|across|boolean|Optional. True to merge cells in each row of the specified range as separate merged cells. The default value is False.|

#### Returns
[Range](range.md)

#### Examples
```js

```
### select()
Selects the specified range in the Excel UI.

#### Syntax
```js
rangeObject.select();
```

#### Parameters
None

#### Returns
void

#### Examples

```js

Excel.run(function (ctx) {
	var sheetName = "Sheet1";
	var rangeAddress = "F5:F10"; 
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.select();
	return ctx.sync(); 
	}); 
});
```

### unmerge()
Separates a merged area into individual cells. If the range contains no merged cell, this API will do no operations on the Range.

#### Syntax
```js
rangeObject.unmerge();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js

```
### Property access examples

Below example uses range address to get the range object.

```js

Excel.run(function (ctx) {
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8"; 
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.load(cellCount);
	return ctx.sync().then(function() {
		Console.log(range.cellCount);
	});
});
```

Below example uses a named-range to get the range object.

```js

Excel.run(function (ctx) { 
	var rangeName = 'MyRange';
	var range = ctx.workbook.names.getItem(rangeName).range;
	range.load(cellCount);
	return ctx.sync().then(function() {
		Console.log(range.cellCount);
	});
});
```

The example below sets number-format, values and formulas on a grid that contains 2x3 grid.

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F5:G7";
	var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
	var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
	var formula = [[null,null], [null,null], [null,"=G6-G5"]];
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.numberFormat = numberFormat;
	range.values = values;
	range.formula = formula;
	range.load(text);
	return ctx.sync().then(function() {
		Console.log(range.text);
	});
});
```
Get the worksheet containing the range. 

```js
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var namedItem = names.getItem('MyRange');
	range = namedItem.range;
	var rangeWorksheet = range.worksheet;
	rangeWorksheet.load(name);
	return ctx.sync().then(function() {
			Console.log(rangeWorksheet.name);
	});
});
```
