# ChartTrendline Object (JavaScript API for Excel)

This object represents the attributes for a chart trendline object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|backward|double|Represents the number of periods that the trendline extends backward.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|displayEquation|bool|True if the equation for the trendline is displayed on the chart.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|displayRSquared|bool|True if the R-squared for the trendline is displayed on the chart.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|forward|double|Represents the number of periods that the trendline extends forward.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|intercept|object|Represents the intercept value of the trendline. Can be set to a numeric value or an empty string (for automatic values). The returned value is always a number.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|movingAveragePeriod|int|Represents the period of a chart trendline, only for trendline with MovingAverage type.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Represents the name of the trendline. Can be set to a string value, or can be set to null value represents automatic values. The returned value is always a string|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|polynomialOrder|int|Represents the order of a chart trendline, only for trendline with Polynomial type.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Represents the type of a chart trendline. Possible values are: Linear, Expontential, Logarithmic, MovingAvg, Polynomial, Power.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartTrendlineFormat](charttrendlineformat.md)|Represents the formatting of a chart trendline. Read-only.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Delete the trendline object.|[Beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Delete the trendline object.

#### Syntax
```js
chartTrendlineObject.delete();
```

#### Parameters
None

#### Returns
void
