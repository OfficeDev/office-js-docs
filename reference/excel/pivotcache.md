# PivotCache Object (JavaScript API for Excel)

Represents the memory cache for a PivotTable report.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|int|Identifier for the specified PivotCache object. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|index|int|Returns a value that represents the index number of the object within the collection of similar objects. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|version|string|Returns the version of Microsoft Excel in which the PivotCache was created. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[refresh()](#refresh)|void|Causes the specified cache to be refreshed and associated chart to be redrawn immediately.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### refresh()
Causes the specified cache to be refreshed and associated chart to be redrawn immediately.

#### Syntax
```js
pivotCacheObject.refresh();
```

#### Parameters
None

#### Returns
void
