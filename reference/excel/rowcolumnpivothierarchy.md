# RowColumnPivotHierarchy Object (JavaScript API for Excel)

Represents the Excel RowColumnPivotHierarchy.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|Id of the RowColumnPivotHierarchy. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Name of the RowColumnPivotHierarchy.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|position|int|Position of the RowColumnPivotHierarchy.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|fields|[PivotFieldCollection](pivotfieldcollection.md)|Returns the PivotFields associated with the RowColumnPivotHierarchy. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[setToDefault()](#settodefault)|void|Reset the RowColumnPivotHierarchy back to its default values.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### setToDefault()
Reset the RowColumnPivotHierarchy back to its default values.

#### Syntax
```js
rowColumnPivotHierarchyObject.setToDefault();
```

#### Parameters
None

#### Returns
void
