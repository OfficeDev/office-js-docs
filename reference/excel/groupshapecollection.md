# GroupShapeCollection Object (JavaScript API for Excel)

Represents a shape collection inside a shape group.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[GroupShape[]](groupshape.md)|A collection of groupShape objects. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Returns the number of shapes in the group shape. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[Shape](shape.md)|Gets a shape using its name.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(shapeId: string)](#getitemshapeid-string)|[Shape](shape.md)|Returns a shape identified by the shape id. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Shape](shape.md)|Gets a shape based on its position in the collection.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Returns the number of shapes in the group shape. Read-only.

#### Syntax
```js
groupShapeCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(name: string)
Gets a shape using its name.

#### Syntax
```js
groupShapeCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the shape to be retrieved.|

#### Returns
[Shape](shape.md)

### getItem(shapeId: string)
Returns a shape identified by the shape id. Read-only.

#### Syntax
```js
groupShapeCollectionObject.getItem(shapeId);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|shapeId|string|The identifier for the shape.|

#### Returns
[Shape](shape.md)

### getItemAt(index: number)
Gets a shape based on its position in the collection.

#### Syntax
```js
groupShapeCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[Shape](shape.md)
