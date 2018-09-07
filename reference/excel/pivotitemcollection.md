# PivotItemCollection Object (JavaScript API for Excel)

Represents a collection of all the Pivot Items related to their parent PivotField.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[PivotItem[]](pivotitem.md)|A collection of pivotItem objects. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Gets the number of pivot hierarchies in the collection.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[PivotItem](pivotitem.md)|Gets a PivotHierarchy by its name or id.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[PivotItem](pivotitem.md)|Gets a PivotHierarchy by name. If the PivotHierarchy does not exist, will return a null object.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Gets the number of pivot hierarchies in the collection.

#### Syntax
```js
pivotItemCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(name: string)
Gets a PivotHierarchy by its name or id.

#### Syntax
```js
pivotItemCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the PivotTable to be retrieved.|

#### Returns
[PivotItem](pivotitem.md)

### getItemOrNullObject(name: string)
Gets a PivotHierarchy by name. If the PivotHierarchy does not exist, will return a null object.

#### Syntax
```js
pivotItemCollectionObject.getItemOrNullObject(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the PivotHierarchy to be retrieved.|

#### Returns
[PivotItem](pivotitem.md)
