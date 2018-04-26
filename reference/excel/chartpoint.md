# ChartPoint Object (JavaScript API for Excel)

Represents a point of a series in a chart.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|hasDataLabel|bool|Represents whether a data point has datalabel. Not applicable for surface charts.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|markerBackgroundColor|string|HTML color code representation of the marker background color of data point. E.g. #FF0000 represents Red.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|markerForegroundColor|string|HTML color code representation of the marker foreground color of data point. E.g. #FF0000 represents Red.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|markerSize|int|Represents marker size of data point.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|markerStyle|string|Represents marker style of a chart data point. Possible values are: Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|value|object|Returns the value of a chart point. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|dataLabel|[ChartDataLabel](chartdatalabel.md)|Returns the data label of a chart point. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|format|[ChartPointFormat](chartpointformat.md)|Encapsulates the format properties chart point. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

