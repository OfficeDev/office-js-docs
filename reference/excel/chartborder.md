# ChartBorder Object (JavaScript API for Excel)

Represents the border formatting for a chart element.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Clear the border color of a chart element.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|Sets the border formatting of a chart element to a uniform color.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### clear()
Clear the border color of a chart element.

#### Syntax
```js
chartBorderObject.clear();
```

#### Parameters
None

#### Returns
void

### setSolidColor(color: string)
Sets the border formatting of a chart element to a uniform color.

#### Syntax
```js
chartBorderObject.setSolidColor(color);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|color|string|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|

#### Returns
void
