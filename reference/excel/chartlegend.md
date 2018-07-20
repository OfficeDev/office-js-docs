# ChartLegend Object (JavaScript API for Excel)

Represents the legend in a chart.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|height|double|Represents the height, in points, of the legend on the chart. Null if legend is not visible.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|left|double|Represents the left, in points, of a chart legend. Null if legend is not visible.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|overlay|bool|Boolean value for whether the chart legend should overlap with the main body of the chart.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|position|string|Represents the position of the legend on the chart. Possible values are: Top, Bottom, Left, Right, Corner, Custom.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showShadow|bool|Represents if the legend has a shadow on the chart.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|top|double|Represents the top of a chart legend.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|A boolean value the represents the visibility of a ChartLegend object.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|width|double|Represents the width, in points, of the legend on the chart. Null if legend is not visible.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartLegendFormat](chartlegendformat.md)|Represents the formatting of a chart legend, which includes fill and font formatting. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|legendEntries|[ChartLegendEntryCollection](chartlegendentrycollection.md)|Represents a collection of legendEntries in the legend. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None


## Method Details

### Property access examples

Get the `position` of Chart Legend from Chart1

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var legend = chart.legend;
	legend.load('position');
	return ctx.sync().then(function() {
			console.log(legend.position);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Set to show legend of Chart1 and make it on top of the chart.

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.legend.visible = true;
	chart.legend.position = "top"; 
	chart.legend.overlay = false; 
	return ctx.sync().then(function() {
			console.log("Legend Shown ");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
``` 
