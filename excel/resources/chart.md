# Chart Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Represents a chart object in a workbook.

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|height|double|Represents the height, in points, of the chart object.|
|left|double|The distance, in points, from the left side of the chart to the worksheet origin.|
|name|string|Represents the name of a chart object.|
|top|double|Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
|width|double|Represents the width, in points, of the chart object.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|axes|[ChartAxes](chartaxes.md)|Represents chart axes. Read-only.|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|Represents the datalabels on the chart. Read-only.|
|format|[ChartAreaFormat](chartareaformat.md)|Encapsulates the format properties for the chart area. Read-only.|
|legend|[ChartLegend](chartlegend.md)|Represents the legend for the chart. Read-only.|
|series|[ChartSeriesCollection](chartseriescollection.md)|Represents either a single series or collection of series in the chart. Read-only.|
|title|[ChartTitle](charttitle.md)|Represents the title of the specified chart, including the text, visibility, position and formating of the title. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the chart object.|
|[getImage(width?: number, height?: number)](#getimagewidth-number-height-number)|string|Return a chart as a base64 string.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[setData(sourceData: Range or string, seriesBy: string)](#setdatasourcedata-range-or-string-seriesby-string)|void|Resets the source data for the chart.|
|[setPosition(startCell: Range or string, endCell: Range or string)](#setpositionstartcell-range-or-string-endcell-range-or-string)|void|Positions the chart relative to cells on the worksheet.|

## Method Details

### delete()
Deletes the chart object.

#### Syntax
```js
chartObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.delete();
	return ctx.sync(); 
}); 
```
### getImage(width?: number, height?: number)
Return a chart as a base64 string.

#### Syntax
```js
chartObject.getImage(width, height);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|width|number|Optional. The width of the output image in pixels. Default is the width of the original chart.|
|height|number|Optional. The height of the output image in pixels. Default is the height of the original chart. |

#### Returns
string

#### Examples
```js

```
### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

#### Examples
```js

```
### setData(sourceData: Range or string, seriesBy: string)
Resets the source data for the chart.

#### Syntax
```js
chartObject.setData(sourceData, seriesBy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|sourceData|Range or string|The address or name of the range that contains the source data. If an address or a worksheet-scoped name is used, it must include the worksheet name (e.g. "Sheet1!A5:B9"). |
|seriesBy|string|Optional. Specifies the way columns or rows are used as data series on the chart. Can be one of the following: Auto (default), Rows, Columns.  Possible values are: Auto, Columns, Rows|

#### Returns
void

#### Examples

Set the `sourceData` to be "A1:B4" and `seriesBy` to be "Columns"

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var sourceData = "A1:B4";
	chart.setData(sourceData, "Columns");
	return ctx.sync(); 
}); 
```

### setPosition(startCell: Range or string, endCell: Range or string)
Positions the chart relative to cells on the worksheet.

#### Syntax
```js
chartObject.setPosition(startCell, endCell);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|startCell|Range or string|The start cell. This is where the chart will be moved to. The start cell is the top-left or top-right cell, depending on the user's right-to-left display settings.|
|endCell|Range or string|Optional. (Optional) The end cell. If specified, the chart's width and height will be set to fully cover up this cell/range.|

#### Returns
void

#### Examples


```js
Excel.run(function (ctx) { 
	var sheetName = "Charts";
	var sourceData = sheetName + "!" + "A1:B4";
	var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("pie", sourceData, "auto");
	chart.width = 500;
	chart.height = 300;
	chart.setPosition("C2", null);
	return ctx.sync(); 
}); 
```

### Property access examples

Get a chart named "Chart1"

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.load('name');
	return ctx.sync().then(function() {
			Console.log(chart.name);
	});
});
```

Update a chart including renaming, positioning and resizing.

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.name="New Name";
	chart.top = 100;
	chart.left = 100;
	chart.height = 200;
	chart.weight = 200;
	return ctx.sync(); 
}); 
```

Rename the chart to new name, resize the chart to 200 points in both height and weight. Move Chart1 to 100 points to the top and left. 

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
	chart.name="New Name";	
	chart.top = 100;
	chart.left = 100;
	chart.height =200;
	chart.width =200;
	return ctx.sync(); 
}); 
```
