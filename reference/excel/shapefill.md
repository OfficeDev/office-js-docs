# ShapeFill Object (JavaScript API for Excel)

Represents the fill formatting for a shape object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|foreColor|string|Represents the shape fill fore color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|transparency|double|Returns or sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). For API not supported shape types  or special fill type with inconsistent transparencies, return null. For example, gradient fill type could have inconsistent transparencies.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|type|string|Returns the fill type of the shape. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Clears the fill formatting of a shape object.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|Sets the fill formatting of a shape object to a uniform color, fill type changeing to Solid Fill.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### clear()
Clears the fill formatting of a shape object.

#### Syntax
```js
shapeFillObject.clear();
```

#### Parameters
None

#### Returns
void

### setSolidColor(color: string)
Sets the fill formatting of a shape object to a uniform color, fill type changeing to Solid Fill.

#### Syntax
```js
shapeFillObject.setSolidColor(color);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|color|string|A string that represents the fill color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|

#### Returns
void
