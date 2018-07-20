# ChartLineFormat Object (JavaScript API for Excel)

Encapsulates the formatting options for line elements.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|color|string|HTML color code representing the color of lines in the chart.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|weight|int|Represents weight of the line, in points.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|lineStyle|string|Represents the line style.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Clear the line format of a chart element.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### clear()
Clear the line format of a chart element.

#### Syntax
```js
chartLineFormatObject.clear();
```

#### Parameters
None

#### Returns
void

#### Examples

Clear the line format of the major gridlines on value axis of the Chart named "Chart1"

```js
Excel.run(function (ctx) { 
	var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueAxis.majorGridlines;	
	gridlines.format.line.clear();
	return ctx.sync().then(function() {
			console.log("Chart Major Gridlines Format Cleared");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### Property access examples

Set chart major gridlines on value axis to be red.

```js
Excel.run(function (ctx) {
	var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueAxis.majorGridlines;
	gridlines.format.line.color = "#FF0000";
	return ctx.sync().then(function () {
		console.log("Chart Gridlines Color Updated");
	});
}).catch(function (error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```
