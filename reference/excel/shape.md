# Shape Object (JavaScript API for Excel)

Represents a generic shape object in the worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|altTextDescription|string|Returns or sets the alternative descriptive text string for a Shape object when the object is saved to a Web page.|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|
|altTextTitle|string|Returns or sets the alternative title text string for a Shape object when the object is saved to a Web page.|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|
|height|double|Represents the height, in points, of the shape.|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|
|id|int|Represents the shape identifier. Read-only.|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|
|left|double|The distance, in points, from the left side of the shape to the left of the worksheet.|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Represents the name of the shape. Read-only.|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|
|rotation|int|Represents the rotation, in degrees, of the shape.|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|
|top|double|The distance, in points, from the top edge of the shape to the top of the worksheet.|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|
|width|double|Represents the width, in points, of the shape.|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|
|zOrderPosition|int|Returns the position of the specified shape in the z-order, the very bottom shape's z-order value is 0. Read-only.|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|geometricShape|[GeometricShape](geometricshape.md)|Returns the geometric shape for the shape object. Error will be thrown, if the shape object is other shape type (Like, Image, SmartArt, etc.) rather than GeometricShape. Read-only.|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|
|image|[Image](image.md)|Returns the image for the shape object. Error will be thrown, if the shape object is other shape type (Like, GeometricShape, SmartArt, etc.) rather than Image. Read-only.|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Returns the type of the specified shape. Read-only.|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[setZOrder(value: string)](#setzordervalue-string)|void|Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).|[ApiSet.InProgressFeatures.ShapeAPIs](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### setZOrder(value: string)
Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).

#### Syntax
```js
shapeObject.setZOrder(value);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|value|string|where to move the specified shape relative to the other shapes.|

#### Returns
void
