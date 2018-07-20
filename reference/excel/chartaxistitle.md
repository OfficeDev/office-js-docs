# ChartAxisTitle Object (JavaScript API for Excel)

Represents the title of a chart axis.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|text|string|Represents the axis title.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|A boolean that specifies the visibility of an axis title.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartAxisTitleFormat](chartaxistitleformat.md)|Represents the formatting of chart axis title. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[setFormula(formula: string)](#setformulaformula-string)|void|A string value that represents the formula of chart axis title using A1-style notation.|[ApiSet.InProgressFeatures.ChartingAPIWave2](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### setFormula(formula: string)
A string value that represents the formula of chart axis title using A1-style notation.

#### Syntax
```js
chartAxisTitleObject.setFormula(formula);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|formula|string| a string that present the formula to set |

#### Returns
void
### Property access examples
Get the `text` of Chart Axis Title from the value axis of Chart1.

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var title = chart.axes.valueAxis.title;
	title.load('text');
	return ctx.sync().then(function() {
			console.log(title.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Add "Values" as the title for the value Axis

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.axes.valueAxis.title.text = "Values";
	return ctx.sync().then(function() {
			console.log("Axis Title Added ");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
