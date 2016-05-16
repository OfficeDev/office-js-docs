# ChartFill Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents the fill formatting for a chart element.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Clear the fill color of a chart element.|1.1|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|Sets the fill formatting of a chart element to a uniform color.|1.1|

## Method Details


### clear()
Clear the fill color of a chart element.

#### Syntax
```js
chartFillObject.clear();
```

#### Parameters
None

#### Returns
void

### setSolidColor(color: string)
Sets the fill formatting of a chart element to a uniform color.

#### Syntax
```js
chartFillObject.setSolidColor(color);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|color|string|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|

#### Returns
void
