# DataPivotHierarchyCollection Object (JavaScript API for Excel)

Represents a collection of DataPivotHierarchy items associated with the PivotTable.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[DataPivotHierarchy[]](datapivothierarchy.md)|A collection of dataPivotHierarchy objects. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(pivotHierarchy: PivotHierarchy)](#addpivothierarchy-pivothierarchy)|[DataPivotHierarchy](datapivothierarchy.md)|Adds the PivotHierarchy to the current axis.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the number of pivot hierarchies in the collection.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[DataPivotHierarchy](datapivothierarchy.md)|Gets a DataPivotHierarchy by its name or id.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[DataPivotHierarchy](datapivothierarchy.md)|Gets a DataPivotHierarchy by name. If the DataPivotHierarchy does not exist, will return a null object.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[remove(DataPivotHierarchy: DataPivotHierarchy)](#removedatapivothierarchy-datapivothierarchy)|void|Removes the PivotHierarchy from the current axis.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(pivotHierarchy: PivotHierarchy)
Adds the PivotHierarchy to the current axis.

#### Syntax
```js
dataPivotHierarchyCollectionObject.add(pivotHierarchy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|pivotHierarchy|PivotHierarchy|TODO|

#### Returns
[DataPivotHierarchy](datapivothierarchy.md)

### getCount()
Gets the number of pivot hierarchies in the collection.

#### Syntax
```js
dataPivotHierarchyCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(name: string)
Gets a DataPivotHierarchy by its name or id.

#### Syntax
```js
dataPivotHierarchyCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the PivotTable to be retrieved.|

#### Returns
[DataPivotHierarchy](datapivothierarchy.md)

### getItemOrNullObject(name: string)
Gets a DataPivotHierarchy by name. If the DataPivotHierarchy does not exist, will return a null object.

#### Syntax
```js
dataPivotHierarchyCollectionObject.getItemOrNullObject(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the DataPivotHierarchy to be retrieved.|

#### Returns
[DataPivotHierarchy](datapivothierarchy.md)

### remove(DataPivotHierarchy: DataPivotHierarchy)
Removes the PivotHierarchy from the current axis.

#### Syntax
```js
dataPivotHierarchyCollectionObject.remove(DataPivotHierarchy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|DataPivotHierarchy|DataPivotHierarchy|TODO|

#### Returns
void
