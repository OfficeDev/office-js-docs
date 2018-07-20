# Worksheet Object (JavaScript API for Excel)

An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|gridlines|bool|Gets or sets the worksheet's gridlines flag.|[ApiSet.InProgressFeatures.LikelyMovedTo18BecauseOfBackCompat](../requirement-sets/excel-api-requirement-sets.md)|
|headings|bool|Gets or sets the worksheet's headings flag.|[ApiSet.InProgressFeatures.LikelyMovedTo18BecauseOfBackCompat](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|The display name of the worksheet.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|position|int|The zero-based position of the worksheet within the workbook.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showGridlines|bool|Gets or sets the worksheet's gridlines flag.|[ApiSet.InProgressFeatures.LikelyMovedTo18BecauseOfBackCompat](../requirement-sets/excel-api-requirement-sets.md)|
|showHeadings|bool|Gets or sets the worksheet's headings flag.|[ApiSet.InProgressFeatures.LikelyMovedTo18BecauseOfBackCompat](../requirement-sets/excel-api-requirement-sets.md)|
|standardHeight|double|Returns the standard (default) height of all the rows in the worksheet, in points. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|standardWidth|double|Returns or sets the standard (default) width of all the columns in the worksheet.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|tabColor|string|Gets or sets the worksheet tab color.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|visibility|string|The Visibility of the worksheet. Possible values are: Visible, Hidden, VeryHidden.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|charts|[ChartCollection](chartcollection.md)|Returns collection of charts that are part of the worksheet. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|freezePanes|[WorksheetFreezePanes](worksheetfreezepanes.md)|Gets an object that can be used to manipulate frozen panes on the worksheet. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|horizontalPageBreaks|[PageBreakCollection](pagebreakcollection.md)|Gets the horizontal page break collection for the worksheet. This collection only contains manual page breaks. Read-only.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|names|[NamedItemCollection](nameditemcollection.md)|Collection of names scoped to the current worksheet. Read-only.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|pageLayout|[PageLayout](pagelayout.md)|Gets the PageLayout object of the worksheet. Read-only.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|pivotTables|[PivotTableCollection](pivottablecollection.md)|Collection of PivotTables that are part of the worksheet. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|protection|[WorksheetProtection](worksheetprotection.md)|Returns sheet protection object for a worksheet. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|shapes|[ShapeCollection](shapecollection.md)|Returns the collection of all the Shape objects on the worksheet. Read-only.|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|
|tables|[TableCollection](tablecollection.md)|Collection of tables that are part of the worksheet. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|verticalPageBreaks|[PageBreakCollection](pagebreakcollection.md)|Gets the vertical page break collection for the worksheet. This collection only contains manual page breaks. Read-only.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|visuals|[VisualCollection](visualcollection.md)|Returns collection of visuals that are part of the worksheet. Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[activate()](#activate)|void|Activate the worksheet in the Excel UI.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[calculate(markAllDirty: bool)](#calculatemarkalldirty-bool)|void|Calculates all cells on a worksheet.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|[copy(positionType: WorksheetPositionType, relativeTo: Worksheet)](#copypositiontype-worksheetpositiontype-relativeto-worksheet)|[Worksheet](worksheet.md)|Copy a worksheet and place it at the specified position. Return the copied worksheet.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|Deletes the worksheet from the workbook.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[findAll(text: string, criteria: WorksheetSearchCriteria)](#findalltext-string-criteria-worksheetsearchcriteria)|[Range](range.md)|Finds all occurances of the given string based on the criteria specified and returns them as a discontiguous range.|[ApiSet.InProgressFeatures.SearchReplace](../requirement-sets/excel-api-requirement-sets.md)|
|[findAllOrNullObject(text: string, criteria: WorksheetSearchCriteria)](#findallornullobjecttext-string-criteria-worksheetsearchcriteria)|[Range](range.md)|Finds all occurances of the given string based on the criteria specified and returns them as a discontiguous range.|[ApiSet.InProgressFeatures.SearchReplace](../requirement-sets/excel-api-requirement-sets.md)|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getNext(visibleOnly: bool)](#getnextvisibleonly-bool)|[Worksheet](worksheet.md)|Gets the worksheet that follows this one. If there are no worksheets following this one, this method will throw an error.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getNextOrNullObject(visibleOnly: bool)](#getnextornullobjectvisibleonly-bool)|[Worksheet](worksheet.md)|Gets the worksheet that follows this one. If there are no worksheets following this one, this method will return a null object.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getPrevious(visibleOnly: bool)](#getpreviousvisibleonly-bool)|[Worksheet](worksheet.md)|Gets the worksheet that precedes this one. If there are no previous worksheets, this method will throw an error.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getPreviousOrNullObject(visibleOnly: bool)](#getpreviousornullobjectvisibleonly-bool)|[Worksheet](worksheet.md)|Gets the worksheet that precedes this one. If there are no previous worksheets, this method will return a null objet.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|Gets the range object specified by the address or name.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)](#getrangebyindexesstartrow-number-startcolumn-number-rowcount-number-columncount-number)|[Range](range.md)|Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRange(valuesOnly: [ApiSet(Version)](#getusedrangevaluesonly-apisetversion)|[Range](range.md)|The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return the top left cell (i.e. it will *not* throw an error).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRangeOrNullObject(valuesOnly: bool)](#getusedrangeornullobjectvaluesonly-bool)|[Range](range.md)|The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return a null object.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[replaceAll(text: string, replacement: string, criteria: ReplaceCriteria)](#replacealltext-string-replacement-string-criteria-replacecriteria)|int|Finds and replaces the given string based on the criteria specified within the current worksheet.|[ApiSet.InProgressFeatures.SearchReplace](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### activate()
Activate the worksheet in the Excel UI.

#### Syntax
```js
worksheetObject.activate();
```

#### Parameters
None

#### Returns
void

#### Examples

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.activate();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### calculate(markAllDirty: bool)
Calculates all cells on a worksheet.

#### Syntax
```js
worksheetObject.calculate(markAllDirty);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|markAllDirty|bool|True, to mark all as dirty.|

#### Returns
void

### copy(positionType: WorksheetPositionType, relativeTo: Worksheet)
Copy a worksheet and place it at the specified position. Return the copied worksheet.

#### Syntax
```js
worksheetObject.copy(positionType, relativeTo);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|positionType|WorksheetPositionType|Optional. Optional.|
|relativeTo|Worksheet|Optional. Optional.|

#### Returns
[Worksheet](worksheet.md)

### delete()
Deletes the worksheet from the workbook.

#### Syntax
```js
worksheetObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### findAll(text: string, criteria: WorksheetSearchCriteria)
Finds all occurances of the given string based on the criteria specified and returns them as a discontiguous range.

#### Syntax
```js
worksheetObject.findAll(text, criteria);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|String to find.|
|criteria|WorksheetSearchCriteria|Additional Criteria.|

#### Returns
[Range](range.md)

### findAllOrNullObject(text: string, criteria: WorksheetSearchCriteria)
Finds all occurances of the given string based on the criteria specified and returns them as a discontiguous range.

#### Syntax
```js
worksheetObject.findAllOrNullObject(text, criteria);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|String to find.|
|criteria|WorksheetSearchCriteria|Additional Criteria.|

#### Returns
[Range](range.md)

### getCell(row: number, column: number)
Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid.

#### Syntax
```js
worksheetObject.getCell(row, column);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|row|number|The row number of the cell to be retrieved. Zero-indexed.|
|column|number|the column number of the cell to be retrieved. Zero-indexed.|

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var cell = worksheet.getCell(0,0);
	cell.load('address');
	return ctx.sync().then(function() {
		console.log(cell.address);
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getNext(visibleOnly: bool)
Gets the worksheet that follows this one. If there are no worksheets following this one, this method will throw an error.

#### Syntax
```js
worksheetObject.getNext(visibleOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|visibleOnly|bool|Optional. Optional. If true, considers only visible worksheets, skipping over any hidden ones.|

#### Returns
[Worksheet](worksheet.md)

### getNextOrNullObject(visibleOnly: bool)
Gets the worksheet that follows this one. If there are no worksheets following this one, this method will return a null object.

#### Syntax
```js
worksheetObject.getNextOrNullObject(visibleOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|visibleOnly|bool|Optional. Optional. If true, considers only visible worksheets, skipping over any hidden ones.|

#### Returns
[Worksheet](worksheet.md)

### getPrevious(visibleOnly: bool)
Gets the worksheet that precedes this one. If there are no previous worksheets, this method will throw an error.

#### Syntax
```js
worksheetObject.getPrevious(visibleOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|visibleOnly|bool|Optional. Optional. If true, considers only visible worksheets, skipping over any hidden ones.|

#### Returns
[Worksheet](worksheet.md)

### getPreviousOrNullObject(visibleOnly: bool)
Gets the worksheet that precedes this one. If there are no previous worksheets, this method will return a null objet.

#### Syntax
```js
worksheetObject.getPreviousOrNullObject(visibleOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|visibleOnly|bool|Optional. Optional. If true, considers only visible worksheets, skipping over any hidden ones.|

#### Returns
[Worksheet](worksheet.md)

### getRange(address: string)
Gets the range object specified by the address or name.

#### Syntax
```js
worksheetObject.getRange(address);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|address|string|Optional. Optional. The address or the name of the range. If not specified, the entire worksheet range is returned.|

#### Returns
[Range](range.md)

#### Examples
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
	var sheetName = "Sheet1";
	var rangeName = 'MyRange';
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeName);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)
Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.

#### Syntax
```js
worksheetObject.getRangeByIndexes(startRow, startColumn, rowCount, columnCount);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|startRow|number|Start row (zero-indexed).|
|startColumn|number|Start column (zero-indexed).|
|rowCount|number|Number of rows to include in the range.|
|columnCount|number|Number of columns to include in the range.|

#### Returns
[Range](range.md)

### getUsedRange(valuesOnly: [ApiSet(Version)
The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return the top left cell (i.e. it will *not* throw an error).

#### Syntax
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|valuesOnly|[ApiSet(Version|Optional. If true, considers only cells with values as used cells (ignoring formatting).|

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	var usedRange = worksheet.getUsedRange();
	usedRange.load('address');
	return ctx.sync().then(function() {
			console.log(usedRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getUsedRangeOrNullObject(valuesOnly: bool)
The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return a null object.

#### Syntax
```js
worksheetObject.getUsedRangeOrNullObject(valuesOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|valuesOnly|bool|Optional. Optional. Considers only cells with values as used cells.|

#### Returns
[Range](range.md)

### replaceAll(text: string, replacement: string, criteria: ReplaceCriteria)
Finds and replaces the given string based on the criteria specified within the current worksheet.

#### Syntax
```js
worksheetObject.replaceAll(text, replacement, criteria);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|String to find.|
|replacement|string|String to replace the original with.|
|criteria|ReplaceCriteria|Additional Replace Criteria.|

#### Returns
int
### Property access examples

Get worksheet properties based on sheet name.

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.load('position')
	return ctx.sync().then(function() {
			console.log(worksheet.position);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Set worksheet position. 

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.position = 2;
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
