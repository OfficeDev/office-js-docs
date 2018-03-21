# PivotField Object (JavaScript API for Excel)

Represents the Excel PivotField.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|Id of the PivotField. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Name of the PivotField.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[PivotItemCollection](pivotitemcollection.md)|Returns the PivotFields associated with the PivotField. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getRange()](#getrange)|[Range](range.md)|Gets the parent range associated with the PivotField.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getRange()
Gets the parent range associated with the PivotField.

#### Syntax
```js
pivotFieldObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)
