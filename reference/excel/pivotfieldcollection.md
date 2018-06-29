# PivotFieldCollection Object (JavaScript API for Excel)

Represents a collection of all the PivotTables that are part of the workbook or worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[PivotField[]](pivotfield.md)|A collection of pivotField objects. Read-only.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Gets the number of pivot hierarchies in the collection.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[PivotField](pivotfield.md)|Gets a PivotHierarchy by its name or id.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[PivotField](pivotfield.md)|Gets a PivotHierarchy by name. If the PivotHierarchy does not exist, will return a null object.|[ApiSet.InProgressFeatures.PivotSharedApis](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Gets the number of pivot hierarchies in the collection.

#### Syntax
```js
pivotFieldCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(name: string)
Gets a PivotHierarchy by its name or id.

#### Syntax
```js
pivotFieldCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the PivotTable to be retrieved.|

#### Returns
[PivotField](pivotfield.md)

### getItemOrNullObject(name: string)
Gets a PivotHierarchy by name. If the PivotHierarchy does not exist, will return a null object.

#### Syntax
```js
pivotFieldCollectionObject.getItemOrNullObject(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the PivotHierarchy to be retrieved.|

#### Returns
[PivotField](pivotfield.md)
