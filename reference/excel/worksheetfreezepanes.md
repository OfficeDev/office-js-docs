# WorksheetFreezePanes Object (JavaScript API for Excel)

Represents the options in sheet protection.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[freezeAt(frozenRange: Range or string)](#freezeatfrozenrange-range-or-string)|void|Sets the frozen cells in the active worksheet view.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|[freezeColumns(count: number)](#freezecolumnscount-number)|void|Freeze the first column(s) of the worksheet in place.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|[freezeRows(count: number)](#freezerowscount-number)|void|Freeze the top row(s) of the worksheet in place.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getLocation()](#getlocation)|[Range](range.md)|Gets a range that describes the frozen cells in the active worksheet view.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getLocationOrNullObject()](#getlocationornullobject)|[Range](range.md)|Gets a range that describes the frozen cells in the active worksheet view.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|[unfreeze()](#unfreeze)|void|Removes all frozen panes in the worksheet.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### freezeAt(frozenRange: Range or string)
Sets the frozen cells in the active worksheet view.

#### Syntax
```js
worksheetFreezePanesObject.freezeAt(frozenRange);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|frozenRange|Range or string|A range that represents the cells to be frozen, or null to remove all frozen panes.|

#### Returns
void

### freezeColumns(count: number)
Freeze the first column(s) of the worksheet in place.

#### Syntax
```js
worksheetFreezePanesObject.freezeColumns(count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|count|number|Optional. Optional number of columns to freeze, or zero to unfreeze all columns|

#### Returns
void

### freezeRows(count: number)
Freeze the top row(s) of the worksheet in place.

#### Syntax
```js
worksheetFreezePanesObject.freezeRows(count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|count|number|Optional. Optional number of rows to freeze, or zero to unfreeze all rows|

#### Returns
void

### getLocation()
Gets a range that describes the frozen cells in the active worksheet view.

#### Syntax
```js
worksheetFreezePanesObject.getLocation();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getLocationOrNullObject()
Gets a range that describes the frozen cells in the active worksheet view.

#### Syntax
```js
worksheetFreezePanesObject.getLocationOrNullObject();
```

#### Parameters
None

#### Returns
[Range](range.md)

### unfreeze()
Removes all frozen panes in the worksheet.

#### Syntax
```js
worksheetFreezePanesObject.unfreeze();
```

#### Parameters
None

#### Returns
void
