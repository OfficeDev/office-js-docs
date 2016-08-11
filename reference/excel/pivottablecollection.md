# PivotTableCollection Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents a collection of all the PivotTables that are part of the workbook or worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[PivotTable[]](pivottable.md)|A collection of pivotTable objects. Read-only.|1.3||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getItem(name: string)](#getitemname-string)|[PivotTable](pivottable.md)|Gets a PivotTable by name.|1.3|
|[getItemOrNull(name: string)](#getitemornullname-string)|[PivotTable](pivottable.md)|Gets a PivotTable by name. If the PivotTable does not exist, the return object's isNull property will be true.|1.3|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|
|[refreshAll()](#refreshall)|void|Refreshes all the PivotTables in the collection.|1.3|

## Method Details


### getItem(name: string)
Gets a PivotTable by name.

#### Syntax
```js
pivotTableCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|name|string|Name of the PivotTable to be retrieved.|

#### Returns
[PivotTable](pivottable.md)

### getItemOrNull(name: string)
Gets a PivotTable by name. If the PivotTable does not exist, the return object's isNull property will be true.

#### Syntax
```js
pivotTableCollectionObject.getItemOrNull(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|name|string|Name of the PivotTable to be retrieved.|

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
Refreshes all the PivotTables in the collection.

#### Syntax
```js
pivotTableCollectionObject.refreshAll();
```

#### Parameters
None

#### Returns
void
