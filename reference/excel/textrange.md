# TextRange Object (JavaScript API for Excel)

Represents the text frame for a shape object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|text|string|Represents the plain text content of the text range.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|font|[ShapeFont](shapefont.md)|Returns a ShapeFont object that represents the font attributes for the text range. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCharacters(start: number, length: number)](#getcharactersstart-number-length-number)|[TextRange](textrange.md)|Returns a TextRange object for characters in the given range.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCharacters(start: number, length: number)
Returns a TextRange object for characters in the given range.

#### Syntax
```js
textRangeObject.getCharacters(start, length);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|start|number|Position of the first character in the returned range.|
|length|number|Optional. Optional. The number of characters to be returned. If omitted, it represents the number of characters from "start" to the end of the last paragraph in TextRange.|

#### Returns
[TextRange](textrange.md)
