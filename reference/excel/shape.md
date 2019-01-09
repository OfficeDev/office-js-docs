# Shape Object (JavaScript API for Excel)

Represents a generic shape object in the worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|altTextDescription|string|Returns or sets the alternative descriptive text string for a Shape object when the object is saved to a Web page.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|altTextTitle|string|Returns or sets the alternative title text string for a Shape object when the object is saved to a Web page.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|height|double|Represents the height, in points, of the shape.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|Represents the shape identifier. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|left|double|The distance, in points, from the left side of the shape to the left of the worksheet.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|level|int|Represents the level of the specified shape. Level 0 means the shape is not part of any group, level 1 means the shape is part of a top-level group, etc. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|lockAspectRatio|bool|Represents if the aspect ratio locked, in boolean, of the shape.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Represents the name of the shape.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|rotation|int|Represents the rotation, in degrees, of the shape.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|top|double|The distance, in points, from the top edge of the shape to the top of the worksheet.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|Represents the visibility, in boolean, of the specified shape.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|width|double|Represents the width, in points, of the shape.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|zOrderPosition|int|Returns the position of the specified shape in the z-order, the very bottom shape's z-order value is 0. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|fill|[ShapeFill](shapefill.md)|Returns the fill formatting of the shape object. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|geometricShape|[GeometricShape](geometricshape.md)|Returns the geometric shape for the shape object. Error will be thrown, if the shape object is other shape type (Like, Image, SmartArt, etc.) rather than GeometricShape. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|geometricShapeType|string|Represents the geometric shape type of the specified shape.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|group|[ShapeGroup](shapegroup.md)|Returns the shape group for the shape object. Error will be thrown, if the shape object is other shape type (Like, Image, SmartArt, etc.) rather than GroupShape. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|image|[Image](image.md)|Returns the image for the shape object. Error will be thrown, if the shape object is other shape type (Like, GeometricShape, SmartArt, etc.) rather than Image. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|line|[Line](line.md)|Returns the line object for the shape object. Error will be thrown, if the shape object is other shape type (Like, GeometricShape, SmartArt, etc.) rather than Image. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|lineFormat|[ShapeLineFormat](shapelineformat.md)|Returns the line formatting of the shape object. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|parentGroup|[Shape](shape.md)|Represents the parent group of the specified shape. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|placement|[Placement](placement.md)|Represents the placment, value that represents the way the object is attached to the cells below it.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|textFrame|[TextFrame](textframe.md)|Returns the textFrame object of a shape. Read only. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Returns the type of the specified shape. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the Shape|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[incrementLeft(increment: double)](#incrementleftincrement-double)|void|Moves the shape horizontally by the specified number of points.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[incrementRotation(increment: double)](#incrementrotationincrement-double)|void|Changes the rotation of the shape around the z-axis by the specified number of degrees.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[incrementTop(increment: double)](#incrementtopincrement-double)|void|Moves the shape vertically by the specified number of points.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[saveAsPicture(format: PictureFormat)](#saveaspictureformat-pictureformat)|string|Saves the shape as a picture and returns the picture in the form of base64 encoded string, using the DPI sets to 96. Only support saves as to Excel.PictureFormat.BMP, Excel.PictureFormat.PNG, Excel.PictureFormat.JPEG and Excel.PictureFormat.GIF.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[scaleHeight(scaleFactor: double, scaleType: ShapeScaleType, scaleFrom: ShapeScaleFrom)](#scaleheightscalefactor-double-scaletype-shapescaletype-scalefrom-shapescalefrom)|void|Scales the height of the shape by a specified factor. For pictures, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[scaleWidth(scaleFactor: double, scaleType: ShapeScaleType, scaleFrom: ShapeScaleFrom)](#scalewidthscalefactor-double-scaletype-shapescaletype-scalefrom-shapescalefrom)|void|Scales the width of the shape by a specified factor. For pictures, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current width.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[setZOrder(value: string)](#setzordervalue-string)|void|Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes the Shape

#### Syntax
```js
shapeObject.delete();
```

#### Parameters
None

#### Returns
void

### incrementLeft(increment: double)
Moves the shape horizontally by the specified number of points.

#### Syntax
```js
shapeObject.incrementLeft(increment);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|increment|double|Specifies how far the shape is to be moved horizontally, in points. A positive value moves the shape to the right; a negative value moves it to the left. If the sheet is RTL, IncrementLeft with a positive value should move the shape to the left instead of right.|

#### Returns
void

### incrementRotation(increment: double)
Changes the rotation of the shape around the z-axis by the specified number of degrees.

#### Syntax
```js
shapeObject.incrementRotation(increment);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|increment|double|Specifies how far the shape is to be rotated horizontally, in degrees. A positive value rotates the shape clockwise; a negative value rotates it counterclockwise.|

#### Returns
void

### incrementTop(increment: double)
Moves the shape vertically by the specified number of points.

#### Syntax
```js
shapeObject.incrementTop(increment);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|increment|double|Specifies how far the shape is to be moved vertically, in points. A positive value moves the shape down; a negative value moves it up.|

#### Returns
void

### saveAsPicture(format: PictureFormat)
Saves the shape as a picture and returns the picture in the form of base64 encoded string, using the DPI sets to 96. Only support saves as to Excel.PictureFormat.BMP, Excel.PictureFormat.PNG, Excel.PictureFormat.JPEG and Excel.PictureFormat.GIF.

#### Syntax
```js
shapeObject.saveAsPicture(format);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|format|PictureFormat|...|

#### Returns
string

### scaleHeight(scaleFactor: double, scaleType: ShapeScaleType, scaleFrom: ShapeScaleFrom)
Scales the height of the shape by a specified factor. For pictures, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.

#### Syntax
```js
shapeObject.scaleHeight(scaleFactor, scaleType, scaleFrom);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|scaleFactor|double|Specifies the ratio between the height of the shape after you resize it and the current or original height.|
|scaleType|ShapeScaleType|OriginalSize to scale the shape relative to its original size. CurrentSize to scale it relative to its current size. You can specify OriginalSize for this argument only if the specified shape is a picture.|
|scaleFrom|ShapeScaleFrom|Optional. Optional. One of the constants of ShapeScaleFrom which specifies which part of the shape retains its position when the shape is scaled. If omitted, it represents the shape's upper left corner retains its position.|

#### Returns
void

### scaleWidth(scaleFactor: double, scaleType: ShapeScaleType, scaleFrom: ShapeScaleFrom)
Scales the width of the shape by a specified factor. For pictures, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current width.

#### Syntax
```js
shapeObject.scaleWidth(scaleFactor, scaleType, scaleFrom);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|scaleFactor|double|Specifies the ratio between the width of the shape after you resize it and the current or original width.|
|scaleType|ShapeScaleType|OriginalSize to scale the shape relative to its original size. CurrentSize to scale it relative to its current size. You can specify OriginalSize for this argument only if the specified shape is a picture.|
|scaleFrom|ShapeScaleFrom|Optional. Optional. One of the constants of ShapeScaleFrom which specifies which part of the shape retains its position when the shape is scaled. If omitted, it represents the shape's upper left corner retains its position.|

#### Returns
void

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
