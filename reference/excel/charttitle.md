# ChartTitle Object (JavaScript API for Excel)

Represents a chart title object of a chart.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|height|double|Returns the height, in points, of the chart title. Read-only. Null if chart title's not visible. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|horizontalAlignment|string|Represents the horizontal alignment for chart title. Possible values are: Center, Left, Right, Justify, Distributed.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|left|double|Represents the distance, in points, from the left edge of chart title to the left edge of chart area. Null if chart title's not visible.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|overlay|bool|Boolean value representing if the chart title will overlay the chart or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showShadow|bool|Represents a boolean value that determines if the chart title has a shadow.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|text|string|Represents the title text of a chart.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|textOrientation|int|Represents the text orientation of chart title. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|top|double|Represents the distance, in points, from the top edge of chart title to the top of chart area. Null if chart title's not visible.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|verticalAlignment|string|Represents the vertical alignment of chart title. Possible values are: Center, Bottom, Top, Justify, Distributed.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|A boolean value the represents the visibility of a chart title object.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|width|double|Returns the width, in points, of the chart title. Read-only. Null if chart title's not visible. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartTitleFormat](charttitleformat.md)|Represents the formatting of a chart title, which includes fill and font formatting. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|position|string|Represents the position of chart title.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getSubstring(start: number, length: number)](#getsubstringstart-number-length-number)|[ChartFormatString](chartformatstring.md)|Get the substring of a chart title. Line break '\n' also counts one charater.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[setFormula(formula: string)](#setformulaformula-string)|void|Sets a string value that represents the formula of chart title using A1-style notation.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getSubstring(start: number, length: number)
Get the substring of a chart title. Line break '\n' also counts one charater.

#### Syntax
```js
chartTitleObject.getSubstring(start, length);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|start|number|Start position of substring to be retrieved. Position start with 0.|
|length|number|Length of substring to be retrieved.|

#### Returns
[ChartFormatString](chartformatstring.md)

### setFormula(formula: string)
Sets a string value that represents the formula of chart title using A1-style notation.

#### Syntax
```js
chartTitleObject.setFormula(formula);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|formula|string| a string that present the formula to set |

#### Returns
void
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
