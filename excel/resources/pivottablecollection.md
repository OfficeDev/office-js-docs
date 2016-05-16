# PivotTableCollection Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents a collection of all the pivot tables that are part of the workbook or worksheet

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Returns the number of pivot tables in the workbook or worksheet. Read-only.|1.3||
|items|[PivotTable[]](pivottable.md)|A collection of pivotTable objects. Read-only.|1.3||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getItem(key: string)](#getitemkey-string)|[PivotTable](pivottable.md)|Gets a pivot table by Name.|1.3|
|[getItemAt(index: number)](#getitematindex-number)|[PivotTable](pivottable.md)|Gets a pivot table based on its position in the collection.|1.3|
|[getItemOrNull(key: string)](#getitemornullkey-string)|[PivotTable](pivottable.md)|Gets a pivot table by Name. If the pivot table does not exist, the return object's isNull property will be true.|1.3|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|
|[refreshAll()](#refreshall)|void|Refreshes all the pivot tables in the collection.|1.3|

## Method Details


### getItem(key: string)
Gets a pivot table by Name.

#### Syntax
```js
pivotTableCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|Name of the pivot table to be retrieved.|

#### Returns
[PivotTable](pivottable.md)

### getItemAt(index: number)
Gets a pivot table based on its position in the collection.

#### Syntax
```js
pivotTableCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[PivotTable](pivottable.md)

### getItemOrNull(key: string)
Gets a pivot table by Name. If the pivot table does not exist, the return object's isNull property will be true.

#### Syntax
```js
pivotTableCollectionObject.getItemOrNull(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|Name of the pivot table to be retrieved.|

#### Returns
[PivotTable](pivottable.md)

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

### refreshAll()
Refreshes all the pivot tables in the collection.

#### Syntax
```js
pivotTableCollectionObject.refreshAll();
```

#### Parameters
None

#### Returns
void
