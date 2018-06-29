# FilterPivotHierarchy Object (JavaScript API for Excel)

Represents the Excel FilterPivotHierarchy.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|enableMultipleFilterItems|bool|Determines whether to allow multiple filter items.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|Id of the FilterPivotHierarchy. Read-only.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Name of the FilterPivotHierarchy.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|position|int|Position of the FilterPivotHierarchy.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|fields|[PivotFieldCollection](pivotfieldcollection.md)|Returns the PivotFields associated with the FilterPivotHierarchy. Read-only.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[setToDefault()](#settodefault)|void|Reset the FilterPivotHierarchy back to its default values.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### setToDefault()
Reset the FilterPivotHierarchy back to its default values.

#### Syntax
```js
filterPivotHierarchyObject.setToDefault();
```

#### Parameters
None

#### Returns
void
