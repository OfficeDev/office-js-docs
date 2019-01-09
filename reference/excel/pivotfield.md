# PivotField Object (JavaScript API for Excel)

Represents the Excel PivotField.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|Id of the PivotField. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Name of the PivotField.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|showAllItems|bool|Determines whether to show all items of the PivotField.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[PivotItemCollection](pivotitemcollection.md)|Returns the PivotFields associated with the PivotField. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|subtotals|[Subtotals](subtotals.md)|Subtotals of the PivotField.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[sortByLabels(sortby: SortBy)](#sortbylabelssortby-sortby)|void|Sorts the PivotField. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the PivotField itself.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[sortByValues(sortby: SortBy, valuesHierarchy: DataPivotHierarchy, pivotItemScope: ()[])](#sortbyvaluessortby-sortby-valueshierarchy-datapivothierarchy-pivotitemscope-)|void|Sorts the PivotField by specified values in a given scope. The scope defines which specific values will be used to sort when|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### sortByLabels(sortby: SortBy)
Sorts the PivotField. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the PivotField itself.

#### Syntax
```js
pivotFieldObject.sortByLabels(sortby);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|sortby|SortBy|Represents whether the sorting is done in an ascending or descending order.|

#### Returns
void

### sortByValues(sortby: SortBy, valuesHierarchy: DataPivotHierarchy, pivotItemScope: ()[])
Sorts the PivotField by specified values in a given scope. The scope defines which specific values will be used to sort when

#### Syntax
```js
pivotFieldObject.sortByValues(sortby, valuesHierarchy, pivotItemScope);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|sortby|SortBy|Represents whether the sorting is done in an ascending or descending order.|
|valuesHierarchy|DataPivotHierarchy|Specifies the values to be used for sorting.|
|pivotItemScope|()[]|Optional. The items that should be used for the scope of the sorting. These will be the
|

#### Returns
void
