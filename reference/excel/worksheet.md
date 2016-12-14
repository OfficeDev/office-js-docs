# Worksheet Object (JavaScript API for Excel)

An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|The display name of the worksheet.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|position|int|The zero-based position of the worksheet within the workbook.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visibility|string|The Visibility of the worksheet. Possible values are: Visible, Hidden, VeryHidden.|[1.1, 1.1 for reading visibility; 1.2 for setting it.](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|charts|[ChartCollection](chartcollection.md)|Returns collection of charts that are part of the worksheet. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|names|[NamedItemCollection](nameditemcollection.md)|Collection of names scoped to the current worksheet. Read-only.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|pivotTables|[PivotTableCollection](pivottablecollection.md)|Collection of PivotTables that are part of the worksheet. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|protection|[WorksheetProtection](worksheetprotection.md)|Returns sheet protection object for a worksheet. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|tables|[TableCollection](tablecollection.md)|Collection of tables that are part of the worksheet. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[activate()](#activate)|void|Activate the worksheet in the Excel UI.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|Deletes the worksheet from the workbook.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getBoundingRange(ranges: ()[])](#getboundingrangeranges-)|[Range](range.md)|Gets the smallest range object that encompasses the provided ranges. For example, the bounding range between "B2:C5" and "D10:E15" is "B2:E16".|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getNextWorksheet(visibleOnly: bool)](#getnextworksheetvisibleonly-bool)|[Worksheet](worksheet.md)|Gets the worksheet that follows this one. If there are no worksheets following this one, this method will throw an error.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getNextWorksheetOrNullObject(visibleOnly: bool)](#getnextworksheetornullobjectvisibleonly-bool)|[Worksheet](worksheet.md)|Gets the worksheet that follows this one. If there are no worksheets following this one, this method will return a null object.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getPreviousWorksheet(visibleOnly: bool)](#getpreviousworksheetvisibleonly-bool)|[Worksheet](worksheet.md)|Gets the worksheet that precedes this one. If there are no previous worksheets, this method will throw an error.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getPreviousWorksheetOrNullObject(visibleOnly: bool)](#getpreviousworksheetornullobjectvisibleonly-bool)|[Worksheet](worksheet.md)|Gets the worksheet that precedes this one. If there are no previous worksheets, this method will return a null objet.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|Gets the range object specified by the address or name.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRangeR1C1(startRow: number, startColumn: number, rowCount: number, columnCount: number)](#getranger1c1startrow-number-startcolumn-number-rowcount-number-columncount-number)|[Range](range.md)|Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[Range](range.md)|The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return the top left cell (i.e.,: it will *not* throw an error).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRangeOrNullObject(valuesOnly: bool)](#getusedrangeornullobjectvaluesonly-bool)|[Range](range.md)|The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return a null object.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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


### getBoundingRange(ranges: ()[])
Gets the smallest range object that encompasses the provided ranges. For example, the bounding range between "B2:C5" and "D10:E15" is "B2:E16".

#### Syntax
```js
worksheetObject.getBoundingRange(ranges);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|ranges|()[]|An array of Range objects or addresses, which will be fully encompassed by the resulting range.|

#### Returns
[Range](range.md)

### getCell(row: number, column: number)
Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid.

#### Syntax
```js
worksheetObject.getCell(row, column);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
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


### getNextWorksheet(visibleOnly: bool)
Gets the worksheet that follows this one. If there are no worksheets following this one, this method will throw an error.

#### Syntax
```js
worksheetObject.getNextWorksheet(visibleOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|visibleOnly|bool|Optional. If true, considers only visible worksheets, skipping over any hidden ones.|

#### Returns
[Worksheet](worksheet.md)

### getNextWorksheetOrNullObject(visibleOnly: bool)
Gets the worksheet that follows this one. If there are no worksheets following this one, this method will return a null object.

#### Syntax
```js
worksheetObject.getNextWorksheetOrNullObject(visibleOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|visibleOnly|bool|Optional. If true, considers only visible worksheets, skipping over any hidden ones.|

#### Returns
[Worksheet](worksheet.md)

### getPreviousWorksheet(visibleOnly: bool)
Gets the worksheet that precedes this one. If there are no previous worksheets, this method will throw an error.

#### Syntax
```js
worksheetObject.getPreviousWorksheet(visibleOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|visibleOnly|bool|Optional. If true, considers only visible worksheets, skipping over any hidden ones.|

#### Returns
[Worksheet](worksheet.md)

### getPreviousWorksheetOrNullObject(visibleOnly: bool)
Gets the worksheet that precedes this one. If there are no previous worksheets, this method will return a null objet.

#### Syntax
```js
worksheetObject.getPreviousWorksheetOrNullObject(visibleOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|visibleOnly|bool|Optional. If true, considers only visible worksheets, skipping over any hidden ones.|

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
|:---------------|:--------|:----------|:---|
|address|string|Optional. The address or the name of the range. If not specified, the entire worksheet range is returned.|

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

### getRangeR1C1(startRow: number, startColumn: number, rowCount: number, columnCount: number)
Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.

#### Syntax
```js
worksheetObject.getRangeR1C1(startRow, startColumn, rowCount, columnCount);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|startRow|number|Start row (zero-indexed).|
|startColumn|number|Start column (zero-indexed).|
|rowCount|number|Number of rows to include in the range.|
|columnCount|number|Number of columns to include in the range.|

#### Returns
[Range](range.md)

### getUsedRange(valuesOnly: bool)
The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return the top left cell (i.e.,: it will *not* throw an error).

#### Syntax
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|valuesOnly|bool|Optional. Considers only cells with values as used cells (ignoring formatting).|

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
|:---------------|:--------|:----------|:---|
|valuesOnly|bool|Considers only cells with values as used cells.|

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
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
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
