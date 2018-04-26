# ChartDataLabels Object (JavaScript API for Excel)

Represents a collection of all the data labels on a chart point.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|autoText|bool|Represents whether data labels automatically generates appropriate text based on context.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|horizontalAlignment|string|Represents the horizontal alignment for chart data label. Possible values are: Center, Left, Right, Justify, Distributed.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|string|Represents the format code for data labels.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormatLinked|bool|Represents whether the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|position|string|DataLabelPosition value that represents the position of the data label. Possible values are: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|separator|string|String representing the separator used for the data labels on a chart.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showBubbleSize|bool|Boolean value representing if the data label bubble size is visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showCategoryName|bool|Boolean value representing if the data label category name is visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showLegendKey|bool|Boolean value representing if the data label legend key is visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showPercentage|bool|Boolean value representing if the data label percentage is visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showSeriesName|bool|Boolean value representing if the data label series name is visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showValue|bool|Boolean value representing if the data label value is visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|textOrientation|int|Represents the text orientation of data labels. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|verticalAlignment|string|Represents the vertical alignment of chart data label. Possible values are: Center, Bottom, Top, Justify, Distributed.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartDataLabelFormat](chartdatalabelformat.md)|Represents the format of chart data labels, which includes fill and font formatting. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None


## Method Details

### Property access examples

Make Series Name shown in Datalabels and set the `position` of datalabels to be "top";

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.datalabels.showValue = true;
	chart.datalabels.position = "top";
	chart.datalabels.showSeriesName = true;
	return ctx.sync().then(function() {
			console.log("Datalabels Shown");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
