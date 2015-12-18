# Worksheet Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.|
|name|string|The display name of the worksheet.|
|position|int|The zero-based position of the worksheet within the workbook.|
|**visibility**|string|The Visibility of the worksheet.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|charts|[ChartCollection](chartcollection.md)|Returns collection of charts that are part of the worksheet. Read-only.|
|**protection**|[WorksheetProtection](worksheetprotection.md)|Returns sheet protection object for a worksheet. Read-only.|
|tables|[TableCollection](tablecollection.md)|Collection of tables that are part of the worksheet. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[activate()](#activate)|void|Activate the worksheet in the Excel UI.|
|[delete()](#delete)|void|Deletes the worksheet from the workbook.|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid.|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|Gets the range object specified by the address or name.|
|**[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)**|[Range](range.md)|The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the worksheet is blank, this function will return the top left cell.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

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

### getCell(row: number, column: number)
Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid.

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

### getRange(address: string)
Gets the range object specified by the address or name.

#### Syntax
```js
worksheetObject.getRange(address);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|address|string|Optional. The address or the name of the range. If not specified, the entire worksheet range is returned.|

#### Returns
[Range](range.md)

### getUsedRange(valuesOnly: bool)
The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the worksheet is blank, this function will return the top left cell.

_**Note**: This is a proposed feature that is still under design phase and hence not yet available as part of the product. The specification is being made available for community review and feedback. The final design may change. Help us make this feature better by providing your feedback [here](https://github.com/OfficeDev/office-js-docs/issues/new?title=ExcelJs-1.2-OpenSpec-range-getusedrange)._

#### Syntax
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|valuesOnly|bool|Optional. Considers only cells with values as used cells (ignores formatting).|

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
