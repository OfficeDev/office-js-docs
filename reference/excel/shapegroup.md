# ShapeGroup Object (JavaScript API for Excel)

Represents a shape group object inside a worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|Represents the shape identifier. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|shape|[Shape](shape.md)|Returns the shape object for the group. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|shapes|[GroupShapeCollection](groupshapecollection.md)|Returns the shape collection in the group. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[ungroup()](#ungroup)|void|Ungroups any grouped shapes in the specified shape group.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### ungroup()
Ungroups any grouped shapes in the specified shape group.

#### Syntax
```js
shapeGroupObject.ungroup();
```

#### Parameters
None

#### Returns
void
