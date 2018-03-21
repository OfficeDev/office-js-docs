# FilterPivotHierarchy Object (JavaScript API for Excel)

Represents the Excel FilterPivotHierarchy.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|enableMultipleFilterItems|bool|Determines whether to allow multiple filter items.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|Id of the FilterPivotHierarchy. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Name of the FilterPivotHierarchy.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|string|Number format of the FilterPivotHierarchy.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|position|int|Position of the FilterPivotHierarchy.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|showAllItems|bool|Determines whether to show all items of the FilterPivotHierarchy.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|currentFilter|[PivotItem](pivotitem.md)|Sets the Filter to the specified PivotItem or returns it, if one is specified.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|fields|[PivotFieldCollection](pivotfieldcollection.md)|Returns the PivotFields associated with the FilterPivotHierarchy. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|subtotals|[Subtotals](subtotals.md)|Subtotals of the FilterPivotHierarchy.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[setToDefault()](#settodefault)|void|Reset the FilterPivotHierarchy back to it's default values.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### setToDefault()
Reset the FilterPivotHierarchy back to it's default values.

#### Syntax
```js
filterPivotHierarchyObject.setToDefault();
```

#### Parameters
None

#### Returns
void
