# ShapeCollection Object (JavaScript API for Excel)

Represents all the shapes in the worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Shape[]](shape.md)|A collection of shape objects. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[addGeometricShape(geometricShapeType: string)](#addgeometricshapegeometricshapetype-string)|[Shape](shape.md)|Adds a geometric shape to worksheet. Returns a Shape object that represents the new shape.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[addGroup(values: ()[])](#addgroupvalues-)|[Shape](shape.md)|Group a subset of shapes in a worksheet. Returns a Shape object that represents the new group of shapes.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[addImage(base64ImageString: string)](#addimagebase64imagestring-string)|[Shape](shape.md)|Creates an image from a base64 string and adds it to worksheet. Returns the Shape object that represents the new Image.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[addLine(startLeft: double, startTop: double, endLeft: double, endTop: double, connectorType: string)](#addlinestartleft-double-starttop-double-endleft-double-endtop-double-connectortype-string)|[Shape](shape.md)|Adds a line to worksheet. Returns a Shape object that represents the new line.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[addSVG(xmlImageString: string)](#addsvgxmlimagestring-string)|[Shape](shape.md)|Creates an SVG from a XML string and adds it to worksheet. Returns a Shape object that represents the new Image.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[addTextBox(text: string)](#addtextboxtext-string)|[Shape](shape.md)|Adds a textbox to worksheet by telling it's text content. Returns a Shape object that represents the new text box.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Returns the number of shapes in the worksheet. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(shapeId: string)](#getitemshapeid-string)|[Shape](shape.md)|Returns a shape identified by the shape id. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[Shape](shape.md)|Gets a shape using its name.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Shape](shape.md)|Gets a shape based on its position in the collection.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### addGeometricShape(geometricShapeType: string)
Adds a geometric shape to worksheet. Returns a Shape object that represents the new shape.

#### Syntax
```js
shapeCollectionObject.addGeometricShape(geometricShapeType);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|geometricShapeType|string|Represents the geometric type of the shape.|

#### Returns
[Shape](shape.md)

### addGroup(values: ()[])
Group a subset of shapes in a worksheet. Returns a Shape object that represents the new group of shapes.

#### Syntax
```js
shapeCollectionObject.addGroup(values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|values|()[]|An array of shape ID or shape objects.|

#### Returns
[Shape](shape.md)

### addImage(base64ImageString: string)
Creates an image from a base64 string and adds it to worksheet. Returns the Shape object that represents the new Image.

#### Syntax
```js
shapeCollectionObject.addImage(base64ImageString);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64ImageString|string|A base64 encoded image in JPEG or PNG formats.|

#### Returns
[Shape](shape.md)

### addLine(startLeft: double, startTop: double, endLeft: double, endTop: double, connectorType: string)
Adds a line to worksheet. Returns a Shape object that represents the new line.

#### Syntax
```js
shapeCollectionObject.addLine(startLeft, startTop, endLeft, endTop, connectorType);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|startLeft|double|The distance, in points, from the start left of the line to the left side of the worksheet.|
|startTop|double|The distance, in points, from the start top of the line to the top of the worksheet.|
|endLeft|double|The distance, in points, from the end left of the line to the left of the worksheet.|
|endTop|double|The distance, in points, from the end top of the line to the top of the worksheet.|
|connectorType|string|Optional. Represents the connector type.|

#### Returns
[Shape](shape.md)

### addSVG(xmlImageString: string)
Creates an SVG from a XML string and adds it to worksheet. Returns a Shape object that represents the new Image.

#### Syntax
```js
shapeCollectionObject.addSVG(xmlImageString);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|xmlImageString|string|An XML string that represents an SVG.|

#### Returns
[Shape](shape.md)

### addTextBox(text: string)
Adds a textbox to worksheet by telling it's text content. Returns a Shape object that represents the new text box.

#### Syntax
```js
shapeCollectionObject.addTextBox(text);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|Optional. Represents the text that will be shown in the created text box.|

#### Returns
[Shape](shape.md)

### getCount()
Returns the number of shapes in the worksheet. Read-only.

#### Syntax
```js
shapeCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(shapeId: string)
Returns a shape identified by the shape id. Read-only.

#### Syntax
```js
shapeCollectionObject.getItem(shapeId);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|shapeId|string|The identifier for the shape.|

#### Returns
[Shape](shape.md)

### getItem(name: string)
Gets a shape using its name.

#### Syntax
```js
shapeCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the shape to be retrieved.|

#### Returns
[Shape](shape.md)

### getItemAt(index: number)
Gets a shape based on its position in the collection.

#### Syntax
```js
shapeCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[Shape](shape.md)
