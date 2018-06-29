# FilterPivotHierarchyCollection Object (JavaScript API for Excel)

Represents a collection of FilterPivotHierarchy items associated with the PivotTable.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[FilterPivotHierarchy[]](filterpivothierarchy.md)|A collection of filterPivotHierarchy objects. Read-only.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(pivotHierarchy: PivotHierarchy)](#addpivothierarchy-pivothierarchy)|[FilterPivotHierarchy](filterpivothierarchy.md)|Adds the PivotHierarchy to the current axis.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the number of pivot hierarchies in the collection.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[FilterPivotHierarchy](filterpivothierarchy.md)|Gets a FilterPivotHierarchy by its name or id.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[FilterPivotHierarchy](filterpivothierarchy.md)|Gets a FilterPivotHierarchy by name. If the FilterPivotHierarchy does not exist, will return a null object.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|[remove(filterPivotHierarchy: FilterPivotHierarchy)](#removefilterpivothierarchy-filterpivothierarchy)|void|Removes the PivotHierarchy from the current axis.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(pivotHierarchy: PivotHierarchy)
Adds the PivotHierarchy to the current axis.

#### Syntax
```js
filterPivotHierarchyCollectionObject.add(pivotHierarchy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|pivotHierarchy|PivotHierarchy|..|

#### Returns
[FilterPivotHierarchy](filterpivothierarchy.md)

### getCount()
Gets the number of pivot hierarchies in the collection.

#### Syntax
```js
filterPivotHierarchyCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(name: string)
Gets a FilterPivotHierarchy by its name or id.

#### Syntax
```js
filterPivotHierarchyCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the PivotTable to be retrieved.|

#### Returns
[FilterPivotHierarchy](filterpivothierarchy.md)

### getItemOrNullObject(name: string)
Gets a FilterPivotHierarchy by name. If the FilterPivotHierarchy does not exist, will return a null object.

#### Syntax
```js
filterPivotHierarchyCollectionObject.getItemOrNullObject(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the FilterPivotHierarchy to be retrieved.|

#### Returns
[FilterPivotHierarchy](filterpivothierarchy.md)

### remove(filterPivotHierarchy: FilterPivotHierarchy)
Removes the PivotHierarchy from the current axis.

#### Syntax
```js
filterPivotHierarchyCollectionObject.remove(filterPivotHierarchy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|filterPivotHierarchy|FilterPivotHierarchy|..|

#### Returns
void
