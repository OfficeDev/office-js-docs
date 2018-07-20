# ChartTrendline Object (JavaScript API for Excel)

This object represents the attributes for a chart trendline object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|backwardPeriod|double|Represents the number of periods that the trendline extends backward.|[ApiSet.InProgressFeatures.ChartingAPIWave2](../requirement-sets/excel-api-requirement-sets.md)|
|forwardPeriod|double|Represents the number of periods that the trendline extends forward.|[ApiSet.InProgressFeatures.ChartingAPIWave2](../requirement-sets/excel-api-requirement-sets.md)|
|intercept|object|Represents the intercept value of the trendline. Can be set to a numeric value or an empty string (for automatic values). The returned value is always a number.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|movingAveragePeriod|int|Represents the period of a chart trendline. Only applicable for trendline with MovingAverage type.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Represents the name of the trendline. Can be set to a string value, or can be set to null value represents automatic values. The returned value is always a string|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|polynomialOrder|int|Represents the order of a chart trendline. Only applicable for trendline with Polynomial type.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|showEquation|bool|True if the equation for the trendline is displayed on the chart.|[ApiSet.InProgressFeatures.ChartingAPIWave2](../requirement-sets/excel-api-requirement-sets.md)|
|showRSquared|bool|True if the R-squared for the trendline is displayed on the chart.|[ApiSet.InProgressFeatures.ChartingAPIWave2](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartTrendlineFormat](charttrendlineformat.md)|Represents the formatting of a chart trendline. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|label|[ChartTrendlineLabel](charttrendlinelabel.md)|Represents the label of a chart trendline. Read-only.|[ApiSet.InProgressFeatures.ChartingAPIWave2](../requirement-sets/excel-api-requirement-sets.md)|
|type|[ChartTrendlineType](charttrendlinetype.md)|Represents the type of a chart trendline.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Delete the trendline object.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

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
