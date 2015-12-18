# Range Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Range represents a set of one or more contiguous cells such as a cell, a row, a column, block of cells, etc.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|address|string|Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. Sheet1!A1:B4). Read-only.|
|addressLocal|string|Represents range reference for the specified range in the language of the user. Read-only.|
|cellCount|int|Number of cells in the range. Read-only.|
|columnCount|int|Represents the total number of columns in the range. Read-only.|
|columnIndex|int|Represents the column number of the first cell in the range. Zero-indexed. Read-only.|
|formulas|object[][]|Represents the formula in A1-style notation.|
|formulasLocal|object[][]|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|
|**formulasR1C1**|object[][]|Represents the formula in R1C1-style notation.|
|numberFormat|object[][]|Represents Excel's number format code for the given cell.|
|rowCount|int|Returns the total number of rows in the range. Read-only.|
|rowIndex|int|Returns the row number of the first cell in the range. Zero-indexed. Read-only.|
|text|object[][]|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|
|valueTypes|string|Represents the type of data of each cell. Read-only.|
|values|object[][]|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|format|[RangeFormat](rangeformat.md)|Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties. Read-only.|
|**sort**|[RangeSort](rangesort.md)|The worksheet containing the current range. Read-only.|
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current range. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
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
|**[getUsedRange(valuesOnly: ApiSetVersion)](#getusedrangevaluesonly-apisetversion)**|[Range](range.md)|Returns the used range of the given range object.|
|[insert(shift: string)](#insertshift-string)|[Range](range.md)|Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|**[merge(across: bool)](#mergeacross-bool)**|void|Merge the range cells into one region in the worksheet.|
|[select()](#select)|void|Selects the specified range in the Excel UI.|
|**[unmerge()](#unmerge)**|void|Unmerge the range cells into separate cells.|

## Method Details


### clear(applyTo: string)
Clear range values, format, fill, border, etc.

#### Syntax
```js
rangeObject.clear(applyTo);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|applyTo|string|Optional. Determines the type of clear action.|

#### Returns
void

### delete(shift: string)
Deletes the cells associated with the range.

#### Syntax
```js
rangeObject.delete(shift);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|shift|string|Specifies which way to shift the cells.|

#### Returns
void

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

### getUsedRange(valuesOnly: [ApiSetVersion)
Returns the used range of the given range object.

_**Note**: This is a proposed feature that is still under design phase and hence not yet available as part of the product. The specification is being made available for community review and feedback. The final design may change. Help us make this feature better by providing your feedback [here](https://github.com/OfficeDev/office-js-docs/issues/new?title=ExcelJs-1.2-OpenSpec-range-getusedrange)._

#### Syntax
```js
rangeObject.getUsedRange(valuesOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|valuesOnly|[ApiSetVersion|Considers only cells with values as used cells.|

#### Returns
[Range](range.md)

### insert(shift: string)
Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.

#### Syntax
```js
rangeObject.insert(shift);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|shift|string|Specifies which way to shift the cells.|

#### Returns
[Range](range.md)

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

### merge(across: bool)
Merge the range cells into one region in the worksheet.

_**Note**: This is a proposed feature that is still under design phase and hence not yet available as part of the product. The specification is being made available for community review and feedback. The final design may change. Help us make this feature better by providing your feedback [here](https://github.com/OfficeDev/office-js-docs/issues/new?title=ExcelJs-1.2-OpenSpec-range-merge)._

#### Syntax
```js
rangeObject.merge(across);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|across|bool|Optional. Set true to merge cells in each row of the specified range as separate merged cells. The default value is false.|

#### Returns
void

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

### unmerge()
Unmerge the range cells into separate cells.

_**Note**: This is a proposed feature that is still under design phase and hence not yet available as part of the product. The specification is being made available for community review and feedback. The final design may change. Help us make this feature better by providing your feedback [here](https://github.com/OfficeDev/office-js-docs/issues/new?title=ExcelJs-1.2-OpenSpec-range-merge)._

#### Syntax
```js
rangeObject.unmerge();
```

#### Parameters
None

#### Returns
void
