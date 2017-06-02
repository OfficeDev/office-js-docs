# ChartTrendline Object (JavaScript API for Excel)

This object represents the attributes for a chart trendline object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|movingAveragePeriod|int|Represents the period of a chart trendline, only for trendline with MovingAverage type.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|polynomialOrder|int|Represents the order of a chart trendline, only for trendline with Polynomial type.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Represents the type of a chart trendline. Possible values are: Linear, Expontential, Logarithmic, MovingAverage, Polynomial, Power.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartTrendlineFormat](charttrendlineformat.md)|Represents the formatting of a chart trendline. Read-only.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

