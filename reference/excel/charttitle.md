# ChartTitle Object (JavaScript API for Excel)

Represents a chart title object of a chart.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|horizontalAlignment|string|Represents the horizontal alignment for the specified object. Possible values are: Center, Left, Right, Justify, Distributed.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|overlay|bool|Boolean value representing if the chart title will overlay the chart or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|text|string|Represents the title text of a chart.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|A boolean value the represents the visibility of a chart title object.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartTitleFormat](charttitleformat.md)|Represents the formatting of a chart title, which includes fill and font formatting. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getSubstring(start: number, length: number)](#getsubstringstart-number-start-number)|[ChartFormatString](chartformatstring.md)|Get the characters of a chart title. Line break '\n' also counts one charater.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getSubstring(start: number, length: number)
Get the characters of a chart title. Line break '\n' also counts one charater.

#### Syntax
```js
chartTitleObject.getSubstring(start, length);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|start|number|The start index of the sub string|
|length|number|The length of the sub string|

#### Returns
[ChartFormatString](chartformatstring.md)
### Property access examples

Get the `text` of Chart Title from Chart1.

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

var title = chart.title;
title.load('text');
return ctx.sync().then(function() {
		console.log(title.text);
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
});
```

Set the `text` of Chart Title to "My Chart" and Make it show on top of the chart without overlaying.

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.title.text= "My Chart"; 
chart.title.visible=true;
chart.title.overlay=true;

return ctx.sync().then(function() {
		console.log("Char Title Changed");
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
});
```
