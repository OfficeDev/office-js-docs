# TextFrame Object (JavaScript API for Excel)

Represents the text frame for a shape object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|bottomMargin|double|Represents the bottom margin, in points, of the text frame.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|hasText|bool|Specifies whether the TextFrame contains text. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|leftMargin|double|Represents the left margin, in points, of the text frame.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|rightMargin|double|Represents the right margin, in points, of the text frame.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|topMargin|double|Represents the top margin, in points, of the text frame.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|autoSize|[ShapeAutoSize](shapeautosize.md)|Gets or sets the auto sizing settings for the text frame. A text frame can be set to auto size the text to fit the text frame, or auto size the text frame to fit the text, or without auto sizing.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|horizontalAlignment|[ShapeTextHorizontalAlignType](shapetexthorizontalaligntype.md)|Represents the horizontal alignment of the text frame.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|horizontalOverflow|[ShapeTextHorzOverflowType](shapetexthorzoverflowtype.md)|Represents the horizontal overflow type of the text frame.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|orientation|[ShapeTextOrientationType](shapetextorientationtype.md)|Represents the text orientation of the text frame.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|readingOrder|[ShapeTextReadingOrder](shapetextreadingorder.md)|Represents the reading order of the text frame, RTL or LTR.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|textRange|[TextRange](textrange.md)|Represents the text range in the text frame.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|verticalAlignment|[ShapeTextVerticalAlignType](shapetextverticalaligntype.md)|Represents the vertical alignment of the text frame.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|verticalOverflow|[ShapeTextVertOverflowType](shapetextvertoverflowtype.md)|Represents the vertical overflow type of the text frame.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[deleteText()](#deletetext)|void|Deletes all the text in the textframe.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### deleteText()
Deletes all the text in the textframe.

#### Syntax
```js
textFrameObject.deleteText();
```

#### Parameters
None

#### Returns
void
