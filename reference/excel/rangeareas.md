# RangeAreas Object (JavaScript API for Excel)

RangeAreas represents a collection of one or more rectangular ranges in the same worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|address|string|Returns the RageAreas reference in A1-style. Address value will contain the worksheet name for each rectangular block of cells (e.g. "Sheet1!A1:B4, Sheet1!D1:D4"). Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|addressLocal|string|Returns the RageAreas reference in the user locale. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|areaCount|int|Returns the number of rectangular ranges that comprise this RangeAreas object. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|cellCount|int|Returns the number of cells in the RangeAreas object, summing up the cell counts of all of the individual rectangular ranges. Returns -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|isEntireColumn|bool|Indicates whether all the ranges on this RangeAreas object represent entire columns (e.g., "A:C, Q:Z"). Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|isEntireRow|bool|Indicates whether all the ranges on this RangeAreas object represent entire rows (e.g., "1:3, 5:7"). Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|style|string|Represents the style for all ranges in this RangeAreas object.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|areas|[RangeCollection](rangecollection.md)|Returns a collection of rectangular ranges that comprise this RangeAreas object. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|conditionalFormats|[ConditionalFormatCollection](conditionalformatcollection.md)|Returns a collection of ConditionalFormats that intersect with any cells in this RangeAreas object. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|dataValidation|[DataValidation](datavalidation.md)|Returns a dataValidation object for all ranges in the RangeAreas. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|format|[RangeFormat](rangeformat.md)|Returns a rangeFormat object, encapsulating the the font, fill, borders, alignment, and other properties for all ranges in the RangeAreas object. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|Returns the worksheet for the current RangeAreas. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[calculate()](#calculate)|void|Calculates all cells in the RangeAreas.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[clear(applyTo: string)](#clearapplyto-string)|void|Clears values, format, fill, border, etc on each of the areas that comprise this RangeAreas object.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[convertDataTypeToText()](#convertdatatypetotext)|void|Converts all cells in the RangeAreas with datatypes into text.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[convertToLinkedDataType(serviceID: number, languageCulture: string)](#converttolinkeddatatypeserviceid-number-languageculture-string)|void|Converts all cells in the RangeAreas into linked datatype.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[copyFrom(sourceRange: Range or RangeAreas or string, copyType: RangeCopyType, skipBlanks: bool, transpose: bool)](#copyfromsourcerange-range-or-rangeareas-or-string-copytype-rangecopytype-skipblanks-bool-transpose-bool)|void|Copies cell data or formatting from the source range or RangeAreas to the current RangeAreas.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireColumn()](#getentirecolumn)|[RangeAreas](rangeareas.md)|Returns a RangeAreas object that represents the entire columns of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11, H2", it returns a RangeAreas that represents columns "B:E, H:H").|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireRow()](#getentirerow)|[RangeAreas](rangeareas.md)|Returns a RangeAreas object that represents the entire rows of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11", it returns a RangeAreas that represents rows "4:11").|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersection(anotherRange: Range or RangeAreas or string)](#getintersectionanotherrange-range-or-rangeareas-or-string)|[RangeAreas](rangeareas.md)|Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, an ItemNotFound error will be thrown.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersectionOrNullObject(anotherRange: Range or RangeAreas or string)](#getintersectionornullobjectanotherrange-range-or-rangeareas-or-string)|[RangeAreas](rangeareas.md)|Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, a null object is returned.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](#getoffsetrangeareasrowoffset-number-columnoffset-number)|[RangeAreas](rangeareas.md)|Returns an RangeAreas object that is shifted by the specific row and column offset. The dimension of the returned RangeAreas will match the original object. If the resulting RangeAreas is forced outside the bounds of the worksheet grid, an error will be thrown.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getSpecialCells(cellType: SpecialCellType, cellValueType: SpecialCellValueType)](#getspecialcellscelltype-specialcelltype-cellvaluetype-specialcellvaluetype)|[RangeAreas](rangeareas.md)|Returns a RangeAreas object that represents all the cells that match the specified type and value. Throws an error if no special cells are found that match the criteria.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getSpecialCellsOrNullObject(cellType: SpecialCellType, cellValueType: SpecialCellValueType)](#getspecialcellsornullobjectcelltype-specialcelltype-cellvaluetype-specialcellvaluetype)|[RangeAreas](rangeareas.md)|Returns a RangeAreas object that represents all the cells that match the specified type and value. Returns a null object if no special cells are found that match the criteria.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getTables(fullyContained: bool)](#gettablesfullycontained-bool)|[TableScopedCollection](tablescopedcollection.md)|Returns a scoped collection of tables that overlap with any range in this RangeAreas object.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRangeAreas(valuesOnly: bool)](#getusedrangeareasvaluesonly-bool)|[RangeAreas](rangeareas.md)|Returns the used RangeAreas that comprises all the used areas of individual rectangular ranges in the RangeAreas object.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRangeAreasOrNullObject(valuesOnly: bool)](#getusedrangeareasornullobjectvaluesonly-bool)|[RangeAreas](rangeareas.md)|Returns the used RangeAreas that comprises all the used areas of individual rectangular ranges in the RangeAreas object.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[setDirty()](#setdirty)|void|Sets the RangeAreas to be recalculated when the next recalculation occurs.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### calculate()
Calculates all cells in the RangeAreas.

#### Syntax
```js
rangeAreasObject.calculate();
```

#### Parameters
None

#### Returns
void

### clear(applyTo: string)
Clears values, format, fill, border, etc on each of the areas that comprise this RangeAreas object.

#### Syntax
```js
rangeAreasObject.clear(applyTo);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|applyTo|string|Optional. Optional. Determines the type of clear action. Possible values are: `All` Default-option,`Formats` ,`Contents` |

#### Returns
void

### convertDataTypeToText()
Converts all cells in the RangeAreas with datatypes into text.

#### Syntax
```js
rangeAreasObject.convertDataTypeToText();
```

#### Parameters
None

#### Returns
void

### convertToLinkedDataType(serviceID: number, languageCulture: string)
Converts all cells in the RangeAreas into linked datatype.

#### Syntax
```js
rangeAreasObject.convertToLinkedDataType(serviceID, languageCulture);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|serviceID|number|The Service ID which will be used to query the data.|
|languageCulture|string|Language Culture to query the service for.|

#### Returns
void

### copyFrom(sourceRange: Range or RangeAreas or string, copyType: RangeCopyType, skipBlanks: bool, transpose: bool)
Copies cell data or formatting from the source range or RangeAreas to the current RangeAreas.

#### Syntax
```js
rangeAreasObject.copyFrom(sourceRange, copyType, skipBlanks, transpose);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|sourceRange|Range or RangeAreas or string|The source range or RangeAreas to copy from. When the source RangeAreas has multiple ranges, it must be in the outline form which can be created by removing full rows or columns from a rectangular range.|
|copyType|RangeCopyType|Optional. The type of cell data or formatting to copy over. Default is "All".|
|skipBlanks|bool|Optional. True if to skip blank cells in the source range or RangeAreas. Default is false.|
|transpose|bool|Optional. True if to transpose the cells in the destination RangeAreas. Default is false.|

#### Returns
void

### getEntireColumn()
Returns a RangeAreas object that represents the entire columns of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11, H2", it returns a RangeAreas that represents columns "B:E, H:H").

#### Syntax
```js
rangeAreasObject.getEntireColumn();
```

#### Parameters
None

#### Returns
[RangeAreas](rangeareas.md)

### getEntireRow()
Returns a RangeAreas object that represents the entire rows of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11", it returns a RangeAreas that represents rows "4:11").

#### Syntax
```js
rangeAreasObject.getEntireRow();
```

#### Parameters
None

#### Returns
[RangeAreas](rangeareas.md)

### getIntersection(anotherRange: Range or RangeAreas or string)
Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, an ItemNotFound error will be thrown.

#### Syntax
```js
rangeAreasObject.getIntersection(anotherRange);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|anotherRange|Range or RangeAreas or string|The range, RangeAreas object or range address that will be used to determine the intersection.|

#### Returns
[RangeAreas](rangeareas.md)

### getIntersectionOrNullObject(anotherRange: Range or RangeAreas or string)
Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, a null object is returned.

#### Syntax
```js
rangeAreasObject.getIntersectionOrNullObject(anotherRange);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|anotherRange|Range or RangeAreas or string|The range, RangeAreas, or address that will be used to determine the intersection.|

#### Returns
[RangeAreas](rangeareas.md)

### getOffsetRangeAreas(rowOffset: number, columnOffset: number)
Returns an RangeAreas object that is shifted by the specific row and column offset. The dimension of the returned RangeAreas will match the original object. If the resulting RangeAreas is forced outside the bounds of the worksheet grid, an error will be thrown.

#### Syntax
```js
rangeAreasObject.getOffsetRangeAreas(rowOffset, columnOffset);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowOffset|number|The number of rows (positive, negative, or 0) by which the RangeAreas is to be offset. Positive values are offset downward, and negative values are offset upward.|
|columnOffset|number|The number of columns (positive, negative, or 0) by which the RangeAreas is to be offset. Positive values are offset to the right, and negative values are offset to the left.|

#### Returns
[RangeAreas](rangeareas.md)

### getSpecialCells(cellType: SpecialCellType, cellValueType: SpecialCellValueType)
Returns a RangeAreas object that represents all the cells that match the specified type and value. Throws an error if no special cells are found that match the criteria.

#### Syntax
```js
rangeAreasObject.getSpecialCells(cellType, cellValueType);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|cellType|SpecialCellType|The type of cells to include.|
|cellValueType|SpecialCellValueType|Optional. If cellType is either Constants or Formulas, this argument is used to determine which types of cells to include in the result. These values can be combined together to return more than one type. The default is to select all constants or formulas, no matter what the type.|

#### Returns
[RangeAreas](rangeareas.md)

### getSpecialCellsOrNullObject(cellType: SpecialCellType, cellValueType: SpecialCellValueType)
Returns a RangeAreas object that represents all the cells that match the specified type and value. Returns a null object if no special cells are found that match the criteria.

#### Syntax
```js
rangeAreasObject.getSpecialCellsOrNullObject(cellType, cellValueType);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|cellType|SpecialCellType|The type of cells to include.|
|cellValueType|SpecialCellValueType|Optional. If cellType is either Constants or Formulas, this argument is used to determine which types of cells to include in the result. These values can be combined together to return more than one type. The default is to select all constants or formulas, no matter what the type.|

#### Returns
[RangeAreas](rangeareas.md)

### getTables(fullyContained: bool)
Returns a scoped collection of tables that overlap with any range in this RangeAreas object.

#### Syntax
```js
rangeAreasObject.getTables(fullyContained);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|fullyContained|bool|Optional. If true, returns only tables that are fully contained within the range bounds. Default is false.|

#### Returns
[TableScopedCollection](tablescopedcollection.md)

### getUsedRangeAreas(valuesOnly: bool)
Returns the used RangeAreas that comprises all the used areas of individual rectangular ranges in the RangeAreas object.

#### Syntax
```js
rangeAreasObject.getUsedRangeAreas(valuesOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|valuesOnly|bool|Optional. Whether to only consider cells with values as used cells. Default is false.|

#### Returns
[RangeAreas](rangeareas.md)

### getUsedRangeAreasOrNullObject(valuesOnly: bool)
Returns the used RangeAreas that comprises all the used areas of individual rectangular ranges in the RangeAreas object.

#### Syntax
```js
rangeAreasObject.getUsedRangeAreasOrNullObject(valuesOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|valuesOnly|bool|Optional. Whether to only consider cells with values as used cells.|

#### Returns
[RangeAreas](rangeareas.md)

### setDirty()
Sets the RangeAreas to be recalculated when the next recalculation occurs.

#### Syntax
```js
rangeAreasObject.setDirty();
```

#### Parameters
None

#### Returns
void
