# PivotItem Object (JavaScript API for Excel)

Represents an item in a PivotTable field.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|calculated|bool|True if the PivotTable item is a calculated field or item. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|drilledDown|bool|True if the flag for the specified PivotTable field or PivotTable item is set to drilled (expanded, or visible).|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Returns or sets a String value representing the name of the object.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|position|int|Returns or sets a Long value that represents the position of the item in its field, if the item is currently showing.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|recordCount|int|Returns the number of records in the PivotTable cache or the number of cache records that contain the specified item. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showDetail|bool|True if the outline is expanded for the specified range (so that the detail of the column or row is visible). The specified range must be a single summary column or row in an outline. For the PivotItem object (or the Range object if the range is in a PivotTable report), this property is set to True if the item is showing detail.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sourceName|string|Returns a value that represents the specified object╬ô├ç├ûs name, as it appears in the original source data, for the specified PivotTable report. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|Returns or sets a Boolean value that determines whether the object is visible.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|pivotField|[PivotField](pivotfield.md)|Returns the parent Pivot Field for the current PivotItem. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getDataRange()](#getdatarange)|[Range](range.md)|Returns a Range object that represents the range of the current PivotItem.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getDataRange()
Returns a Range object that represents the range of the current PivotItem.

#### Syntax
```js
pivotItemObject.getDataRange();
```

#### Parameters
None

#### Returns
[Range](range.md)
