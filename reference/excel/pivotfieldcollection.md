# PivotFieldCollection Object (JavaScript API for Excel)

A collection of all the PivotField objects in a PivotTable report.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[PivotField[]](pivotfield.md)|A collection of pivotField objects. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Returns a value that represents the number of objects in the collection.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(nameOrIndex: object)](#getitemnameorindex-object)|[PivotField](pivotfield.md)|Returns a single object from a collection based on name or index.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[PivotField](pivotfield.md)|Returns a single object from a collection based on the index.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Returns a value that represents the number of objects in the collection.

#### Syntax
```js
pivotFieldCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(nameOrIndex: object)
Returns a single object from a collection based on name or index.

#### Syntax
```js
pivotFieldCollectionObject.getItem(nameOrIndex);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|nameOrIndex|object|Name or index of the item to be retrieved.|

#### Returns
[PivotField](pivotfield.md)

### getItemAt(index: number)
Returns a single object from a collection based on the index.

#### Syntax
```js
pivotFieldCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index of the item. Zero-indexed.|

#### Returns
[PivotField](pivotfield.md)
