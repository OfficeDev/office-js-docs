# SlicerItemCollection Object (JavaScript API for Excel)

Represents a collection of all the slicer item objects on the slicer.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[SlicerItem[]](sliceritem.md)|A collection of slicerItem objects. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Returns the number of slicer items in the slicer.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[SlicerItem](sliceritem.md)|Gets a slicer item object using its key or name.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[SlicerItem](sliceritem.md)|Gets a slicer item based on its position in the collection.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[SlicerItem](sliceritem.md)|Gets a slicer item using its key or name. If the slicer item does not exist, will return a null object.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Returns the number of slicer items in the slicer.

#### Syntax
```js
slicerItemCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(key: string)
Gets a slicer item object using its key or name.

#### Syntax
```js
slicerItemCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|The key or name of the slicer item.|

#### Returns
[SlicerItem](sliceritem.md)

### getItemAt(index: number)
Gets a slicer item based on its position in the collection.

#### Syntax
```js
slicerItemCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[SlicerItem](sliceritem.md)

### getItemOrNullObject(key: string)
Gets a slicer item using its key or name. If the slicer item does not exist, will return a null object.

#### Syntax
```js
slicerItemCollectionObject.getItemOrNullObject(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|Key or name of the slicer to be retrieved.|

#### Returns
[SlicerItem](sliceritem.md)
