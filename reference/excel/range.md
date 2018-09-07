# Range Object (JavaScript API for Excel)

Range represents a set of one or more contiguous cells such as a cell, a row, a column, block of cells, etc.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|address|string|Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. "Sheet1!A1:B4"). Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|addressLocal|string|Represents range reference for the specified range in the language of the user. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|cellCount|int|Number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnCount|int|Represents the total number of columns in the range. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnHidden|bool|Represents if all columns of the current range are hidden.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|columnIndex|int|Represents the column number of the first cell in the range. Zero-indexed. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulas|object[][]|Represents the formula in A1-style notation.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulasLocal|object[][]|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulasR1C1|object[][]|Represents the formula in R1C1-style notation.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|hidden|bool|Represents if all cells of the current range are hidden. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|isEntireColumn|bool|Represents if the current range is an entire column. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|isEntireRow|bool|Represents if the current range is an entire row. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|object[][]|Represents Excel's number format code for the given range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormatLocal|object[][]|Represents Excel's number format code for the given range as a string in the language of the user.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|Returns the total number of rows in the range. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowHidden|bool|Represents if all rows of the current range are hidden.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|rowIndex|int|Returns the row number of the first cell in the range. Zero-indexed. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|style|string|Represents the style of the current range.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|text|string|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|valueTypes|string|Represents the type of data of each cell. Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|conditionalFormats|[ConditionalFormatCollection](conditionalformatcollection.md)|Collection of ConditionalFormats that intersect the range. Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|dataValidation|[DataValidation](datavalidation.md)|Returns a data validation object. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|format|[RangeFormat](rangeformat.md)|Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|hyperlink|[RangeHyperlink](rangehyperlink.md)|Represents the hyperlink for the current range.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|linkedDataTypeState|[LinkedDataTypeState](linkeddatatypestate.md)|Represents the data type state of each cell. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|sort|[RangeSort](rangesort.md)|Represents the range sort of the current range. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current range. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[calculate()](#calculate)|void|Calculates a range of cells on a worksheet.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|[clear(applyTo: string)](#clearapplyto-string)|void|Clear range values, format, fill, border, etc.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[convertDataTypeToText()](#convertdatatypetotext)|void|Converts the range cells with datatypes into text.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[convertToLinkedDataType(serviceID: number, languageCulture: string)](#converttolinkeddatatypeserviceid-number-languageculture-string)|void|Converts the range cells into linked datatype in the worksheet.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[copyFrom(sourceRange: Range or RangeAreas or string, copyType: RangeCopyType, skipBlanks: bool, transpose: bool)](#copyfromsourcerange-range-or-rangeareas-or-string-copytype-rangecopytype-skipblanks-bool-transpose-bool)|void|Copies cell data or formatting from the source range or RangeAreas to the current range.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[delete(shift: string)](#deleteshift-string)|void|Deletes the cells associated with the range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[find(text: string, criteria: SearchCriteria)](#findtext-string-criteria-searchcriteria)|[Range](range.md)|Finds the given string based on the criteria specified.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[findOrNullObject(text: string, criteria: SearchCriteria)](#findornullobjecttext-string-criteria-searchcriteria)|[Range](range.md)|Finds the given string based on the criteria specified.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getAbsoluteResizedRange(numRows: number, numColumns: number)](#getabsoluteresizedrangenumrows-number-numcolumns-number)|[Range](range.md)|Gets a Range object with the same top-left cell as the current Range object, but with the specified numbers of rows and columns.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[getBoundingRect(anotherRange: Range or string)](#getboundingrectanotherrange-range-or-string)|[Range](range.md)|Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E15".|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumn(column: number)](#getcolumncolumn-number)|[Range](range.md)|Gets a column contained in the range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnsAfter(count: number)](#getcolumnsaftercount-number)|[Range](range.md)|Gets a certain number of columns to the right of the current Range object.|[ApiSet.PolyfillableDownTo1_1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnsBefore(count: number)](#getcolumnsbeforecount-number)|[Range](range.md)|Gets a certain number of columns to the left of the current Range object.|[ApiSet.PolyfillableDownTo1_1](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireColumn()](#getentirecolumn)|[Range](range.md)|Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", its `getEntireColumn` is a range that represents columns "B:E").|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireRow()](#getentirerow)|[Range](range.md)|Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", its `GetEntireRow` is a range that represents rows "4:11").|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getImage()](#getimage)|[System.IO.Stream](system.io.stream.md)|Renders the range as a base64-encoded png image.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersection(anotherRange: Range or string)](#getintersectionanotherrange-range-or-string)|[Range](range.md)|Gets the range object that represents the rectangular intersection of the given ranges.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersectionOrNullObject(anotherRange: Range or string)](#getintersectionornullobjectanotherrange-range-or-string)|[Range](range.md)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastCell()](#getlastcell)|[Range](range.md)|Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastColumn()](#getlastcolumn)|[Range](range.md)|Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastRow()](#getlastrow)|[Range](range.md)|Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[Range](range.md)|Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an error will be thrown.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getResizedRange(deltaRows: number, deltaColumns: number)](#getresizedrangedeltarows-number-deltacolumns-number)|[Range](range.md)|Gets a Range object similar to the current Range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.|[ApiSet.PolyfillableDownTo1_1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRow(row: number)](#getrowrow-number)|[Range](range.md)|Gets a row contained in the range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowsAbove(count: number)](#getrowsabovecount-number)|[Range](range.md)|Gets a certain number of rows above the current Range object.|[ApiSet.PolyfillableDownTo1_1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowsBelow(count: number)](#getrowsbelowcount-number)|[Range](range.md)|Gets a certain number of rows below the current Range object.|[ApiSet.PolyfillableDownTo1_1](../requirement-sets/excel-api-requirement-sets.md)|
|[getSpecialCells(cellType: SpecialCellType, cellValueType: SpecialCellValueType)](#getspecialcellscelltype-specialcelltype-cellvaluetype-specialcellvaluetype)|[RangeAreas](rangeareas.md)|Gets the RangeAreas object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getSpecialCellsOrNullObject(cellType: SpecialCellType, cellValueType: SpecialCellValueType)](#getspecialcellsornullobjectcelltype-specialcelltype-cellvaluetype-specialcellvaluetype)|[RangeAreas](rangeareas.md)|Gets the RangeAreas object, comprising one or more ranges, that represents all the cells that match the specified type and value.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getSurroundingRegion()](#getsurroundingregion)|[Range](range.md)|Returns a Range object that represents the surrounding region for the top-left cell in this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[getTables(fullyContained: bool)](#gettablesfullycontained-bool)|[TableScopedCollection](tablescopedcollection.md)|Gets a scoped collection of tables that overlap with the range.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRange(valuesOnly: [ApiSet(Version)](#getusedrangevaluesonly-apisetversion)|[Range](range.md)|Returns the used range of the given range object. If there are no used cells within the range, this function will throw an ItemNotFound error.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRangeOrNullObject(valuesOnly: bool)](#getusedrangeornullobjectvaluesonly-bool)|[Range](range.md)|Returns the used range of the given range object. If there are no used cells within the range, this function will return a null object.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getVisibleView()](#getvisibleview)|[RangeView](rangeview.md)|Represents the visible rows of the current range.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[insert(shift: string)](#insertshift-string)|[Range](range.md)|Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[merge(across: bool)](#mergeacross-bool)|void|Merge the range cells into one region in the worksheet.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[removeDuplicates(columns: int[], includesHeader: bool)](#removeduplicatescolumns-int-includesheader-bool)|[RemoveDuplicatesResult](removeduplicatesresult.md)|Removes duplicate values from the range specified by the columns.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[replaceAll(text: string, replacement: string, criteria: ReplaceCriteria)](#replacealltext-string-replacement-string-criteria-replacecriteria)|int|Finds and replaces the given string based on the criteria specified within the current range.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[select()](#select)|void|Selects the specified range in the Excel UI.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[setDirty()](#setdirty)|void|Set a range to be recalculated when the next recalculation occurs.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[showCard()](#showcard)|void|Displays the card for an active cell if it has rich value content.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[unmerge()](#unmerge)|void|Unmerge the range cells into separate cells.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### calculate()
Calculates a range of cells on a worksheet.

#### Syntax
```js
rangeObject.calculate();
```

#### Parameters
None

#### Returns
void

### clear(applyTo: string)
Clear range values, format, fill, border, etc.

#### Syntax
```js
rangeObject.clear(applyTo);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|applyTo|string|Optional. Optional. Determines the type of clear action. Possible values are: `All` Default-option,`Formats` ,`Contents` |

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
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### convertDataTypeToText()
Converts the range cells with datatypes into text.

#### Syntax
```js
rangeObject.convertDataTypeToText();
```

#### Parameters
None

#### Returns
void

### convertToLinkedDataType(serviceID: number, languageCulture: string)
Converts the range cells into linked datatype in the worksheet.

#### Syntax
```js
rangeObject.convertToLinkedDataType(serviceID, languageCulture);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|serviceID|number|The Service ID which will be used to query the data.|
|languageCulture|string|Language Culture to query the service for.|

#### Returns
void

### copyFrom(sourceRange: Range or RangeAreas or string, copyType: RangeCopyType, skipBlanks: bool, transpose: bool)
Copies cell data or formatting from the source range or RangeAreas to the current range.

#### Syntax
```js
rangeObject.copyFrom(sourceRange, copyType, skipBlanks, transpose);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|sourceRange|Range or RangeAreas or string|The source range or RangeAreas to copy from. When the source RangeAreas has multiple ranges, it must in the outline form which can be created by removing full rows or columns from a rectangular range.|
|copyType|RangeCopyType|Optional. The type of cell data or formatting to copy over. Default is "All".|
|skipBlanks|bool|Optional. True if to skip blank cells in the source range. Default is false.|
|transpose|bool|Optional. True if to transpose the cells in the destination range. Default is false.|

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
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### find(text: string, criteria: SearchCriteria)
Finds the given string based on the criteria specified.

#### Syntax
```js
rangeObject.find(text, criteria);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|String to find.|
|criteria|SearchCriteria|Additional Criteria.|

#### Returns
[Range](range.md)

### findOrNullObject(text: string, criteria: SearchCriteria)
Finds the given string based on the criteria specified.

#### Syntax
```js
rangeObject.findOrNullObject(text, criteria);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|String to find.|
|criteria|SearchCriteria|Additional Criteria.|

#### Returns
[Range](range.md)

### getAbsoluteResizedRange(numRows: number, numColumns: number)
Gets a Range object with the same top-left cell as the current Range object, but with the specified numbers of rows and columns.

#### Syntax
```js
rangeObject.getAbsoluteResizedRange(numRows, numColumns);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|numRows|number|The number of rows of the new range size.|
|numColumns|number|The number of columns of the new range size.|

#### Returns
[Range](range.md)

### getBoundingRect(anotherRange: Range or string)
Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E15".

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
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // Prints Sheet1!D4:H8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getCell(row: number, column: number)
Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.

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
	cell.load('address');
	return ctx.sync().then(function() {
		console.log(cell.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
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
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!B1:B8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getColumnsAfter(count: number)
Gets a certain number of columns to the right of the current Range object.

#### Syntax
```js
rangeObject.getColumnsAfter(count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|count|number|Optional. Optional. The number of columns to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.|

#### Returns
[Range](range.md)

### getColumnsBefore(count: number)
Gets a certain number of columns to the left of the current Range object.

#### Syntax
```js
rangeObject.getColumnsBefore(count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|count|number|Optional. Optional. The number of columns to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.|

#### Returns
[Range](range.md)

### getEntireColumn()
Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", its `getEntireColumn` is a range that represents columns "B:E").

#### Syntax
```js
rangeObject.getEntireColumn();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

Note: the grid properties of the Range (values, numberFormat, formulas) contains `null` since the Range in question is unbounded.

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var rangeEC = range.getEntireColumn();
	rangeEC.load('address');
	return ctx.sync().then(function() {
		console.log(rangeEC.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getEntireRow()
Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", its `GetEntireRow` is a range that represents rows "4:11").

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
	rangeER.load('address');
	return ctx.sync().then(function() {
		console.log(rangeER.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
The grid properties of the Range (values, numberFormat, formulas) contains `null` since the Range in question is unbounded.


### getImage()
Renders the range as a base64-encoded png image.

#### Syntax
```js
rangeObject.getImage();
```

#### Parameters
None

#### Returns
[System.IO.Stream](system.io.stream.md)

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
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!D4:F6
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getIntersectionOrNullObject(anotherRange: Range or string)
Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.

#### Syntax
```js
rangeObject.getIntersectionOrNullObject(anotherRange);
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

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastCell();
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!F8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
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
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!F1:F8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
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
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!A8:F8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```



### getOffsetRange(rowOffset: number, columnOffset: number)
Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an error will be thrown.

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
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!H3:K5
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getResizedRange(deltaRows: number, deltaColumns: number)
Gets a Range object similar to the current Range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.

#### Syntax
```js
rangeObject.getResizedRange(deltaRows, deltaColumns);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|deltaRows|number|The number of rows by which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.|
|deltaColumns|number|The number of columns by which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.|

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

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getRow(1);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!A2:F2
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getRowsAbove(count: number)
Gets a certain number of rows above the current Range object.

#### Syntax
```js
rangeObject.getRowsAbove(count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|count|number|Optional. Optional. The number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.|

#### Returns
[Range](range.md)

### getRowsBelow(count: number)
Gets a certain number of rows below the current Range object.

#### Syntax
```js
rangeObject.getRowsBelow(count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|count|number|Optional. Optional. The number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.|

#### Returns
[Range](range.md)

### getSpecialCells(cellType: SpecialCellType, cellValueType: SpecialCellValueType)
Gets the RangeAreas object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value.

#### Syntax
```js
rangeObject.getSpecialCells(cellType, cellValueType);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|cellType|SpecialCellType|The type of cells to include.|
|cellValueType|SpecialCellValueType|Optional. If cellType is either Constants or Formulas, this argument is used to determine which types of cells to include in the result. These values can be combined together to return more than one type. The default is to select all constants or formulas, no matter what the type.|

#### Returns
[RangeAreas](rangeareas.md)

### getSpecialCellsOrNullObject(cellType: SpecialCellType, cellValueType: SpecialCellValueType)
Gets the RangeAreas object, comprising one or more ranges, that represents all the cells that match the specified type and value.

#### Syntax
```js
rangeObject.getSpecialCellsOrNullObject(cellType, cellValueType);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|cellType|SpecialCellType|The type of cells to include.|
|cellValueType|SpecialCellValueType|Optional. If cellType is either Constants or Formulas, this argument is used to determine which types of cells to include in the result. These values can be combined together to return more than one type. The default is to select all constants or formulas, no matter what the type.|

#### Returns
[RangeAreas](rangeareas.md)

### getSurroundingRegion()
Returns a Range object that represents the surrounding region for the top-left cell in this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.

#### Syntax
```js
rangeObject.getSurroundingRegion();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getTables(fullyContained: bool)
Gets a scoped collection of tables that overlap with the range.

#### Syntax
```js
rangeObject.getTables(fullyContained);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|fullyContained|bool|Optional. If true, returns only tables that are fully contained within the range bounds. The default value is false.|

#### Returns
[TableScopedCollection](tablescopedcollection.md)

### getUsedRange(valuesOnly: [ApiSet(Version)
Returns the used range of the given range object. If there are no used cells within the range, this function will throw an ItemNotFound error.

#### Syntax
```js
rangeObject.getUsedRange(valuesOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|valuesOnly|[ApiSet(Version|Considers only cells with values as used cells.|

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var rangeUR = range.getUsedRange();
	rangeUR.load('address');
	return ctx.sync().then(function() {
		console.log(rangeUR.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getUsedRangeOrNullObject(valuesOnly: bool)
Returns the used range of the given range object. If there are no used cells within the range, this function will return a null object.

#### Syntax
```js
rangeObject.getUsedRangeOrNullObject(valuesOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|valuesOnly|bool|Optional. Considers only cells with values as used cells.|

#### Returns
[Range](range.md)

### getVisibleView()
Represents the visible rows of the current range.

#### Syntax
```js
rangeObject.getVisibleView();
```

#### Parameters
None

#### Returns
[RangeView](rangeview.md)

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
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### merge(across: bool)
Merge the range cells into one region in the worksheet.

#### Syntax
```js
rangeObject.merge(across);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|across|bool|Optional. Optional. Set true to merge cells in each row of the specified range as separate merged cells. The default value is false.|

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:C3";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.merge(true);
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```



#### Examples
```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:C3";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.unmerge();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### removeDuplicates(columns: int[], includesHeader: bool)
Removes duplicate values from the range specified by the columns.

#### Syntax
```js
rangeObject.removeDuplicates(columns, includesHeader);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|columns|int[]|The columns inside the range that may contain duplicates. At least one column needs to be specified. Zero-indexed.|
|includesHeader|bool|True if the input data contains header. Default is false.|

#### Returns
[RemoveDuplicatesResult](removeduplicatesresult.md)

### replaceAll(text: string, replacement: string, criteria: ReplaceCriteria)
Finds and replaces the given string based on the criteria specified within the current range.

#### Syntax
```js
rangeObject.replaceAll(text, replacement, criteria);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|String to find.|
|replacement|string|String to replace the original with.|
|criteria|ReplaceCriteria|Additional Replace Criteria.|

#### Returns
int

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
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### setDirty()
Set a range to be recalculated when the next recalculation occurs.

#### Syntax
```js
rangeObject.setDirty();
```

#### Parameters
None

#### Returns
void

### showCard()
Displays the card for an active cell if it has rich value content.

#### Syntax
```js
rangeObject.showCard();
```

#### Parameters
None

#### Returns
void

### unmerge()
Unmerge the range cells into separate cells.

#### Syntax
```js
rangeObject.unmerge();
```

#### Parameters
None

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:C3";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.unmerge();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### Property access examples

Below example uses range address to get the range object.

```js

Excel.run(function (ctx) {
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8"; 
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.load('cellCount');
	return ctx.sync().then(function() {
		console.log(range.cellCount);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Below example uses a named-range to get the range object.

```js

Excel.run(function (ctx) { 
	var rangeName = 'MyRange';
	var range = ctx.workbook.names.getItem(rangeName).range;
	range.load('cellCount');
	return ctx.sync().then(function() {
		console.log(range.cellCount);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

The example below sets number-format, values and formulas on a grid that contains 2x3 grid.

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F5:G7";
	var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
	var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
	var formulas = [[null,null], [null,null], [null,"=G6-G5"]];
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.numberFormat = numberFormat;
	range.values = values;
	range.formulas= formulas;
	range.load('text');
	return ctx.sync().then(function() {
		console.log(range.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
Get the worksheet containing the range. 

```js
/* This might be broken still - it was broken before because it 
	it was missing 'var', but might still be wrong because of
	getting information without loading properly. */
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var namedItem = names.getItem('MyRange');
	var range = namedItem.range;
	var rangeWorksheet = range.worksheet;
	rangeWorksheet.load('name');
	return ctx.sync().then(function() {
			console.log(rangeWorksheet.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

