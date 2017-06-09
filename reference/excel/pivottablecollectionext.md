# PivotTableCollectionExt Object (JavaScript API for Excel)

A collection of all the PivotTable objects in the specified workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[PivotTable[]](pivottable.md)|A collection of pivotTable objects. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(name: string, address: Range, pivotCache: PivotCache)](#addname-string-address-range-pivotcache-pivotcache)|[PivotTable](pivottable.md)|Adds a new PivotTable report. Returns a PivotTable object.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(name: string, address: Range, pivotCache: PivotCache)
Adds a new PivotTable report. Returns a PivotTable object.

#### Syntax
```js
pivotTableCollectionExtObject.add(name, address, pivotCache);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|The name of the new PivotTable report.|
|address|Range|The cell in the upper-left corner of the PivotTable report's destination range (the range on the worksheet where the resulting report will be placed). You must specify a destination range on the worksheet that contains the PivotTables object specified by expression.|
|pivotCache|PivotCache|The PivotTable cache on which the new PivotTable report is based. The cache provides data for the report.|

#### Returns
[PivotTable](pivottable.md)
