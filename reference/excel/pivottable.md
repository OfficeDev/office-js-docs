# PivotTable Object (JavaScript API for Excel)

Represents an Excel PivotTable.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|Id of the PivotTable. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Name of the PivotTable.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|useCustomSortLists|bool|True if the PivotTable should use custom lists when sorting.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|columnHierarchies|[RowColumnPivotHierarchyCollection](rowcolumnpivothierarchycollection.md)|The Column Pivot Hierarchies of the PivotTable. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|dataHierarchies|[DataPivotHierarchyCollection](datapivothierarchycollection.md)|The Data Pivot Hierarchies of the PivotTable. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|filterHierarchies|[FilterPivotHierarchyCollection](filterpivothierarchycollection.md)|The Filter Pivot Hierarchies of the PivotTable. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|hierarchies|[PivotHierarchyCollection](pivothierarchycollection.md)|The Pivot Hierarchies of the PivotTable. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|layout|[PivotLayout](pivotlayout.md)|The PivotLayout describing the layout and visual structure of the PivotTable. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|rowHierarchies|[RowColumnPivotHierarchyCollection](rowcolumnpivothierarchycollection.md)|The Row Pivot Hierarchies of the PivotTable. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current PivotTable. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the PivotTable.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[refresh()](#refresh)|void|Refreshes the PivotTable.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes the PivotTable.

#### Syntax
```js
pivotTableObject.delete();
```

#### Parameters
None

#### Returns
void

### refresh()
Refreshes the PivotTable.

#### Syntax
```js
pivotTableObject.refresh();
```

#### Parameters
None

#### Returns
void
