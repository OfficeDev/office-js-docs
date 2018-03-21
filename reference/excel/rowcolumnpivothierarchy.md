# RowColumnPivotHierarchy Object (JavaScript API for Excel)

Represents the Excel RowColumnPivotHierarchy.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|Id of the RowColumnPivotHierarchy. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Name of the RowColumnPivotHierarchy.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|string|Number format of the RowColumnPivotHierarchy.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|position|int|Position of the RowColumnPivotHierarchy.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|showAllItems|bool|Determines whether to show all items of the RowColumnPivotHierarchy.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|fields|[PivotFieldCollection](pivotfieldcollection.md)|Returns the PivotFields associated with the RowColumnPivotHierarchy. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|subtotals|[Subtotals](subtotals.md)|Subtotals of the RowColumnPivotHierarchy.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[setToDefault()](#settodefault)|void|Reset the RowColumnPivotHierarchy back to it's default values.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[setTopBottomValueFilter(criteriontype: PivotFilterTopBottomCriterion, rank: number, dataPivotHierarchy: DataPivotHierarchy)](#settopbottomvaluefiltercriteriontype-pivotfiltertopbottomcriterion-rank-number-datapivothierarchy-datapivothierarchy)|void|TopBottom Filter the RowColumnPivotHierarchy.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[sort(ascending: bool, dataPivotHierarchy: DataPivotHierarchy)](#sortascending-bool-datapivothierarchy-datapivothierarchy)|void|Sort the RowColumnPivotHierarchy. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the RowColumnPivotHierarchy itself.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### setToDefault()
Reset the RowColumnPivotHierarchy back to it's default values.

#### Syntax
```js
rowColumnPivotHierarchyObject.setToDefault();
```

#### Parameters
None

#### Returns
void

### setTopBottomValueFilter(criteriontype: PivotFilterTopBottomCriterion, rank: number, dataPivotHierarchy: DataPivotHierarchy)
TopBottom Filter the RowColumnPivotHierarchy.

#### Syntax
```js
rowColumnPivotHierarchyObject.setTopBottomValueFilter(criteriontype, rank, dataPivotHierarchy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|criteriontype|PivotFilterTopBottomCriterion|The criterion to use for the filter.|
|rank|number|The rank of the filter. For percent this has to be between 1 and 100.|
|dataPivotHierarchy|DataPivotHierarchy|The DataPivotHierarchy on which the filter is based on.|

#### Returns
void

### sort(ascending: bool, dataPivotHierarchy: DataPivotHierarchy)
Sort the RowColumnPivotHierarchy. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the RowColumnPivotHierarchy itself.

#### Syntax
```js
rowColumnPivotHierarchyObject.sort(ascending, dataPivotHierarchy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|ascending|bool|Represents whether the sorting is done in an ascending or descending order.|
|dataPivotHierarchy|DataPivotHierarchy|Optional. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the RowColumnPivotHierarchy itself.|

#### Returns
void
