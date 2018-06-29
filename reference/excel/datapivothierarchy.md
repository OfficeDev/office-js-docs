# DataPivotHierarchy Object (JavaScript API for Excel)

Represents the Excel DataPivotHierarchy.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|Id of the DataPivotHierarchy. Read-only.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Name of the DataPivotHierarchy.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|string|Number format of the DataPivotHierarchy.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|position|int|Position of the DataPivotHierarchy.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|field|[PivotField](pivotfield.md)|Returns the PivotFields associated with the DataPivotHierarchy. Read-only.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|showAs|[ShowAsRule](showasrule.md)|Determines whether the data should be sown as a specific summary calculation or not.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|summarizeBy|[AggregationFunction](aggregationfunction.md)|Determines whether to show all items of the DataPivotHierarchy.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[setToDefault()](#settodefault)|void|Reset the DataPivotHierarchy back to its default values.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### setToDefault()
Reset the DataPivotHierarchy back to its default values.

#### Syntax
```js
dataPivotHierarchyObject.setToDefault();
```

#### Parameters
None

#### Returns
void
