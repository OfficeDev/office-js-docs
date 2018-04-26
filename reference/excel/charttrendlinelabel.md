# ChartTrendlineLabel Object (JavaScript API for Excel)

This object represents the attributes for a chart trendline lable object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|autoText|bool|Boolean value representing if trendline label automatically generates appropriate text based on context.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|formula|string|String value that represents the formula of chart trendline label using A1-style notation.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|height|double|Returns the height, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|horizontalAlignment|string|Represents the horizontal alignment for chart trendline label. Possible values are: Center, Left, Right, Justify, Distributed.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|left|double|Represents the distance, in points, from the left edge of chart trendline label to the left edge of chart area. Null if chart trendline label is not visible.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|string|String value that represents the format code for trendline label.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormatLinked|bool|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|text|string|String representing the text of the trendline label on a chart.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|textOrientation|int|Represents the text orientation of chart trendline label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|top|double|Represents the distance, in points, from the top edge of chart trendline label to the top of chart area. Null if chart trendline label is not visible.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|verticalAlignment|string|Represents the vertical alignment of chart trendline label. Possible values are: Center, Bottom, Top, Justify, Distributed.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|width|double|Returns the width, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartTrendlineLabelFormat](charttrendlinelabelformat.md)|Represents the format of chart trendline label. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

