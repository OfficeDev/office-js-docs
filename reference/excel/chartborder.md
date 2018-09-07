# ChartBorder Object (JavaScript API for Excel)

Represents the border formatting of a chart element.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|color|string|HTML color code representing the color of borders in the chart.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|weight|int|Represents weight of the border, in points.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|lineStyle|string|Represents the line style of the border.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Clear the border format of a chart element.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### clear()
Clear the border format of a chart element.

#### Syntax
```js
chartBorderObject.clear();
```

#### Parameters
None

#### Returns
void
