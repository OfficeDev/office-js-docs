# RowColumnPivotHierarchyCollection Object (JavaScript API for Excel)

Represents a collection of RowColumnPivotHierarchy items associated with the PivotTable.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[RowColumnPivotHierarchy[]](rowcolumnpivothierarchy.md)|A collection of rowColumnPivotHierarchy objects. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(pivotHierarchy: PivotHierarchy)](#addpivothierarchy-pivothierarchy)|[RowColumnPivotHierarchy](rowcolumnpivothierarchy.md)|Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column,|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the number of pivot hierarchies in the collection.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[RowColumnPivotHierarchy](rowcolumnpivothierarchy.md)|Gets a RowColumnPivotHierarchy by its name or id.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[RowColumnPivotHierarchy](rowcolumnpivothierarchy.md)|Gets a RowColumnPivotHierarchy by name. If the RowColumnPivotHierarchy does not exist, will return a null object.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)](#removerowcolumnpivothierarchy-rowcolumnpivothierarchy)|void|Removes the PivotHierarchy from the current axis.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(pivotHierarchy: PivotHierarchy)
Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column,

#### Syntax
```js
rowColumnPivotHierarchyCollectionObject.add(pivotHierarchy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|pivotHierarchy|PivotHierarchy|...|

#### Returns
[RowColumnPivotHierarchy](rowcolumnpivothierarchy.md)

### getCount()
Gets the number of pivot hierarchies in the collection.

#### Syntax
```js
rowColumnPivotHierarchyCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(name: string)
Gets a RowColumnPivotHierarchy by its name or id.

#### Syntax
```js
rowColumnPivotHierarchyCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the PivotTable to be retrieved.|

#### Returns
[RowColumnPivotHierarchy](rowcolumnpivothierarchy.md)

### getItemOrNullObject(name: string)
Gets a RowColumnPivotHierarchy by name. If the RowColumnPivotHierarchy does not exist, will return a null object.

#### Syntax
```js
rowColumnPivotHierarchyCollectionObject.getItemOrNullObject(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the RowColumnPivotHierarchy to be retrieved.|

#### Returns
[RowColumnPivotHierarchy](rowcolumnpivothierarchy.md)

### remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)
Removes the PivotHierarchy from the current axis.

#### Syntax
```js
rowColumnPivotHierarchyCollectionObject.remove(rowColumnPivotHierarchy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowColumnPivotHierarchy|RowColumnPivotHierarchy|...|

#### Returns
void
