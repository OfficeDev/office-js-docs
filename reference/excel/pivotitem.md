# PivotItem Object (JavaScript API for Excel)

Represents the Excel PivotItem.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|Id of the PivotItem. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Name of the PivotItem.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|showDetail|bool|Determines whether details are shown or not.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|Determines whether the PivotItem is visible or not.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getRange()](#getrange)|[Range](range.md)|Returns the range of the PivotItem.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getRange()
Returns the range of the PivotItem.

#### Syntax
```js
pivotItemObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)
