# ChartDataLabel Object (JavaScript API for Excel)

Represents the data label of a chart point.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|autoText|bool|Boolean value representing if data label automatically generates appropriate text based on context.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|formula|string|String value that represents the formula of chart data label using A1-style notation.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|height|double|Returns the height, in points, of the chart data label. Read-only. Null if chart data label is not visible. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|horizontalAlignment|string|Represents the horizontal alignment for chart data label. Possible values are: Center, Left, Right, Justify, Distributed.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|left|double|Represents the distance, in points, from the left edge of chart data label to the left edge of chart area. Null if chart data label is not visible.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|string|String value that represents the format code for data label.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormatLinked|bool|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|position|string|DataLabelPosition value that represents the position of the data label. Possible values are: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|separator|string|String representing the separator used for the data label on a chart.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|showBubbleSize|bool|Boolean value representing if the data label bubble size is visible or not.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|showCategoryName|bool|Boolean value representing if the data label category name is visible or not.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|showLegendKey|bool|Boolean value representing if the data label legend key is visible or not.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|showPercentage|bool|Boolean value representing if the data label percentage is visible or not.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|showSeriesName|bool|Boolean value representing if the data label series name is visible or not.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|showValue|bool|Boolean value representing if the data label value is visible or not.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|text|string|String representing the text of the data label on a chart.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|textOrientation|int|Represents the text orientation of chart data label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|top|double|Represents the distance, in points, from the top edge of chart data label to the top of chart area. Null if chart data label is not visible.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|verticalAlignment|string|Represents the vertical alignment of chart data label. Possible values are: Center, Bottom, Top, Justify, Distributed.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|width|double|Returns the width, in points, of the chart data label. Read-only. Null if chart data label is not visible. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartDataLabelFormat](chartdatalabelformat.md)|Represents the format of chart data label. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

