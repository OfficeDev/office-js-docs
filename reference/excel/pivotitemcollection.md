# PivotItemCollection Object (JavaScript API for Excel)

A collection of all the PivotItem objects in a PivotTable field.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[PivotItem[]](pivotitem.md)|A collection of pivotItem objects. Read-only.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Returns a value that represents the number of objects in the collection.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(nameOrIndex: object)](#getitemnameorindex-object)|[PivotItem](pivotitem.md)|Returns a single object from a collection based on name or index.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[PivotItem](pivotitem.md)|Returns a single object from a collection based on the index.|[Pivot1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Returns a value that represents the number of objects in the collection.

#### Syntax
```js
pivotItemCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(nameOrIndex: object)
Returns a single object from a collection based on name or index.

#### Syntax
```js
pivotItemCollectionObject.getItem(nameOrIndex);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|nameOrIndex|object|Name or index of the item to be retrieved.|

#### Returns
[PivotItem](pivotitem.md)

### getItemAt(index: number)
Returns a single object from a collection based on the index.

#### Syntax
```js
pivotItemCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index of the item. Zero-indexed.|

#### Returns
[PivotItem](pivotitem.md)
