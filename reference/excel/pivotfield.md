# PivotField Object (JavaScript API for Excel)

Represents the Excel PivotField.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|Id of the PivotField. Read-only.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Name of the PivotField.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|showAllItems|bool|Determines whether to show all items of the PivotField.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[PivotItemCollection](pivotitemcollection.md)|Returns the PivotFields associated with the PivotField. Read-only.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|subtotals|[Subtotals](subtotals.md)|Subtotals of the PivotField.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[sortByLabels(ascending: bool)](#sortbylabelsascending-bool)|void|Sorts the PivotField. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the PivotField itself.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### sortByLabels(ascending: bool)
Sorts the PivotField. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the PivotField itself.

#### Syntax
```js
pivotFieldObject.sortByLabels(ascending);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|ascending|bool|Represents whether the sorting is done in an ascending or descending order.|

#### Returns
void
