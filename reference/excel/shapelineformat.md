# ShapeLineFormat Object (JavaScript API for Excel)

Represents the line formatting for the shape object. For picture and geometric shape, line formatting represents the border of shape object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|color|string|Represents the line color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|transparency|double|Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has mixed line transparency property (e.g. group type of shape).|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|Represents whether the line formatting of a shape element is visible. Returns null when the shape has mixed line visible property (e.g. group type of shape).|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|weight|int|Represents weight of the line, in points. Returns null when the line is not visible or has mixed line weight property (e.g. group type of shape).|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|dashStyle|string|Represents the line style of the shape. Returns null when line is not visible or has mixed line dash style property (e.g. group type of shape).|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|style|string|Represents the line style of the shape object. Returns null when line is not visible or has mixed line visible property (e.g. group type of shape).|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

